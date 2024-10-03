using System.Formats.Asn1;
using System.Globalization;
using static PulseRecord.Class.CSVRecord;
using CsvHelper;
using PulseRecord.Class;
using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using Microsoft.Extensions.Configuration;
using Topshelf.Configurators;

namespace PulseRecord
{
    public class Worker : BackgroundService
    {
        // Ruta del archivo CSV
        //string filePath = @"C:\Users\R105\Documents\_Desarrollos\Bard\Servicio Instron, crear archivos CSV\Recursos\Actual\CSV\45322545_28.is_tens_Results.csv";
        private readonly ILogger<Worker> _logger;
        private readonly IConfiguration _configuration;
        private readonly TimeSpan _delay;

        public Worker(ILogger<Worker> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
            _delay = TimeSpan.FromSeconds(_configuration.GetValue<int>("WorkerSettings:IntervalInSeconds"));

        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            // Obtener rutas desde appconfig.json
            string directoryPath = _configuration.GetValue<string>("Paths:DirectorioOrigen");//configuration["Paths:DirectorioOrigen"];
            string archivedPath = _configuration.GetValue<string>("Paths:DirectorioDestino");//configuration["Paths:DirectorioDestino"];
            string errorsPath = _configuration.GetValue<string>("Paths:DirectorioError");//configuration["Paths:DirectorioError"];

            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("Aplicación Pulse corriendo a las: {time}", DateTimeOffset.Now);

                // Crear carpetas si no existen
                Directory.CreateDirectory(archivedPath);
                Directory.CreateDirectory(errorsPath);

                // Obtener todos los archivos CSV en el directorio
                var csvFiles = Directory.GetFiles(directoryPath, "*.csv");

                var recipients = EmailService.GetEmailRecipients();

                foreach (var filePath in csvFiles)
                {
                    try
                    {

                        using (var reader = new StreamReader(filePath))
                        {
                              var headerRecord = await Task.Run(() => ReadHeaderAndResults(reader, recipients), stoppingToken);

                            Console.WriteLine($"Procesando archivo: {Path.GetFileName(filePath)}");

                            LogEvent("INFO", "Archivo procesado correctamente", Path.GetFileName(filePath) + " | #Parte: " + headerRecord.NumeroDeParte + " | Lote: " + headerRecord.Lote);

                            // Procesar e imprimir cada registro
                            foreach (var result in headerRecord.Results)
                            {
                                Console.WriteLine($"{headerRecord.CalDate}\t{headerRecord.Equipo}\t{headerRecord.Inspector}\t{headerRecord.Lado}\t{headerRecord.Lote}\t{headerRecord.NumeroDeParte}\t{headerRecord.SelloPel}\t{result.Result}\t{result.MaximumLoad}");
                            }

                            // Imprimir una línea vacía o un separador para indicar el fin del archivo
                            Console.WriteLine(new string('-', 50));
                        }


                        // Mover el archivo a la carpeta "Archived" si no hay errores
                        MoveFile(filePath, archivedPath);
                    }
                    catch (Exception ex)
                    {
                        // Mover el archivo a la carpeta "Errors" si ocurre algún error
                        LogEvent("ERROR", $"Error al procesar el archivo: {ex.Message}", Path.GetFileName(filePath));
                        Console.WriteLine($"Error al procesar el archivo {filePath}: {ex.Message}");
                        MoveFile(filePath, errorsPath);
                    }
                }


                await Task.Delay(_delay, stoppingToken);
            }

        }


        static void MoveFile(string filePath, string destinationFolder)
        {
            string fileName = Path.GetFileName(filePath);
            string destinationPath = Path.Combine(destinationFolder, fileName);

            // Mover el archivo
            File.Move(filePath, destinationPath);
            Console.WriteLine($"Archivo movido a {destinationPath}");
        }



        public CSVRecord ReadHeaderAndResults(StreamReader reader, string recipient)
        {
            Boolean SiFalla = false;
            var record = new CSVRecord();
            var CSVheader = GetHeaderMappings();
            string line;

            // Leer el encabezado
            while ((line = reader.ReadLine()) != null)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue; // Saltar líneas vacías

                // Cuando llegamos a la línea de "Results Table 1", sabemos que la cabecera ha terminado.
                if (line.Contains("Results Table 1"))
                    break;

                var parts = line.Split(new[] { ":," }, StringSplitOptions.None);

                if (parts.Length < 2) continue;
                string header = parts[0].Trim();
                string value = parts[1].Trim('"');

                // Verifica si el encabezado existe en el mapeo
                if (CSVheader.ContainsKey(header))
                {
                    string propertyName = CSVheader[header];

                    // Usar reflexión para asignar el valor al campo correspondiente
                    var property = typeof(CSVRecord).GetProperty(propertyName);
                    if (property != null)
                    {
                        property.SetValue(record, value);
                    }
                }

                //switch (parts[0].Trim())
                //{
                //    case "Cal. Date":
                //        record.CalDate = parts[1].Trim('"');
                //        break;
                //    case "Equipo":
                //        record.Equipo = parts[1].Trim('"');
                //        break;
                //    case "Inspector":
                //        record.Inspector = parts[1].Trim('"');
                //        break;
                //    case "Lado":
                //        record.Lado = parts[1].Trim('"');
                //        break;
                //    case "Lote":
                //        record.Lote = parts[1].Trim('"');
                //        break;
                //    case "Numero De Parte":
                //        record.NumeroDeParte = parts[1].Trim('"');
                //        break;
                //    case "Sello Pel":
                //        record.SelloPel = parts[1].Trim('"');
                //        break;
                //}
            }
            var minMaxValores = ObtenerMinMaxValores(record.NumeroDeParte);

            // Leer la tabla de resultados
            var results = new List<ResultRecord>();
            while ((line = reader.ReadLine()) != null)
            {
                if (string.IsNullOrWhiteSpace(line) || line.Contains("Maximum Load") || line.Contains("(lbf)"))
                    continue; // Saltar líneas vacías o líneas de encabezado

                var parts = line.Split(',');
                if (parts.Length < 2) continue;

                var resultRecord = new ResultRecord
                {
                    Result = int.Parse(parts[0].Trim()),
                    MaximumLoad = double.Parse(parts[1].Trim('"'))
                };

               

                if (resultRecord.MaximumLoad < minMaxValores.Minimo || resultRecord.MaximumLoad > minMaxValores.Maximo)
                {
                    resultRecord.Estado = "Falla";
                }
                else
                {
                    resultRecord.Estado = "Paso";
                    SiFalla = true;
                }

                results.Add(resultRecord);
            }


            // Validar si alguna propiedad no tiene valor
            if (!ValidateRecord(record))
            {
                Console.WriteLine("Falta información en el CSV. Creando Log...");
                LogEvent("INFO", "Faltan datos en el archivo CSV. El archivo CSV no contiene todos los campos requeridos.", " | #Parte: " + record.NumeroDeParte + " | Lote: " + record.Lote);
            }

            // Agregar los resultados al record
            record.Results = results;

            SaveValues(record);

            if (SiFalla)
            {
                //Validar falta de lista de distribucion
                var emailSubject = "PulseRecord App - Notificación de pruebas con fallas";
                if(recipient is not "")
                {
                    EmailService.SendEmailWithCsvData(recipient, emailSubject, results, record);
                }
                else
                {
                    LogEvent("INFO", "No se encontraron correos en lista de distribución. No es posible enviar correo.", " | #Parte: " + record.NumeroDeParte + " | Lote: " + record.Lote);
                }

                
            }
            

            return record;
        }

        public void SaveValues(CSVRecord RS)
        {
            // Configurar el cargador de configuración para leer desde appsettings.json
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            IConfiguration config = builder.Build();

            // Leer la cadena de conexión
            string connectionString = config.GetConnectionString("API");

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand("SP_PulseRecord", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Option", "Register_Instron");
                    cmd.Parameters.AddWithValue("@FileName", RS.NumeroDeParte + '-' + RS.Lote);
                    cmd.Parameters.AddWithValue("@CalDate", RS.CalDate);
                    cmd.Parameters.AddWithValue("@Equipo", RS.Equipo);
                    cmd.Parameters.AddWithValue("@Inspector", RS.Inspector);
                    cmd.Parameters.AddWithValue("@Lado", RS.Lado);
                    cmd.Parameters.AddWithValue("@Lote", RS.Lote );
                    cmd.Parameters.AddWithValue("@NumeroDeParte", RS.NumeroDeParte);
                    cmd.Parameters.AddWithValue("@SelloPel", RS.SelloPel);


                    var resultsTable = new DataTable();
                    resultsTable.Columns.Add("Result", typeof(string));
                    resultsTable.Columns.Add("MaximumLoad", typeof(decimal));
                    resultsTable.Columns.Add("Estado", typeof(string));

                    foreach (var result in RS.Results)
                    {
                        resultsTable.Rows.Add(result.Result, result.MaximumLoad,result.Estado);
                    }


                    var tvpParam = cmd.Parameters.AddWithValue("@Resultado", resultsTable);
                    tvpParam.SqlDbType = SqlDbType.Structured;
                    tvpParam.TypeName = "a_InstronResultsType";

                    // Execute the stored procedure
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void LogEvent(string eventType, string message, string fileName)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            IConfiguration config = builder.Build();

            // Leer la cadena de conexión
            string connectionString = config.GetConnectionString("API");

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand("INSERT INTO a_EventLog (EventType, Message, FileName) VALUES (@EventType, @Message, @FileName)", conn))
                {
                    cmd.Parameters.AddWithValue("@EventType", eventType);
                    cmd.Parameters.AddWithValue("@Message", message);
                    cmd.Parameters.AddWithValue("@FileName", fileName);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public MinMaxValores ObtenerMinMaxValores(string numeroDeParte)
        {
            var minMaxValores = new MinMaxValores();

            string query = "SP_PulseRecord";

            // Configurar el cargador de configuración para leer desde appsettings.json
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            IConfiguration config = builder.Build();

            // Leer la cadena de conexión
            string connectionString = config.GetConnectionString("API");

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Option", "Obtener_limites");
                    cmd.Parameters.AddWithValue("@NumeroDeParte", numeroDeParte);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            minMaxValores.Minimo = Convert.ToDouble(reader.GetDecimal(reader.GetOrdinal("Minimo")));
                            minMaxValores.Maximo = Convert.ToDouble(reader.GetDecimal(reader.GetOrdinal("Maximo")));
                        }
                    }
                }
            }

            return minMaxValores;
        }

        public static Dictionary<string, string> GetHeaderMappings()
        {
            var mappings = new Dictionary<string, string>();
            string query = "SP_PulseRecord";

            // Configurar el cargador de configuración para leer desde appsettings.json
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            IConfiguration config = builder.Build();

            // Leer la cadena de conexión
            string connectionString = config.GetConnectionString("API");

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Option", "Obtener_CSVEncabezados");

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string header = reader["HeaderName"].ToString();
                            string property = reader["PropertyName"].ToString();
                            mappings[header] = property;
                        }
                    }
                }
            }

            return mappings;
        }

        public bool ValidateRecord(CSVRecord record)
        {
            var properties = typeof(CSVRecord).GetProperties();

            foreach (var prop in properties)
            {
                // Obtiene el valor de la propiedad
                var value = prop.GetValue(record)?.ToString();

                // Verifica si la propiedad es nula o está vacía
                if (string.IsNullOrWhiteSpace(value))
                {
                    Console.WriteLine($"La propiedad {prop.Name} no tiene un valor asignado.");
                    return false; // Si alguna propiedad no tiene valor, devuelve falso
                }
            }

            return true; // Si todas las propiedades tienen valor, devuelve verdadero
        }
    }
}