using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices.JavaScript;

namespace PulseRecord.Class
{

    public class EmailService
    {
        public static string GetEmailRecipients()
        {
            string recipients = "";
            // Configurar el cargador de configuración para leer desde appsettings.json
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            IConfiguration config = builder.Build();

            // Leer la cadena de conexión
            string connectionString = config.GetConnectionString("API");

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (var command = new SqlCommand("SP_PulseRecord", conn))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@Option", "Obtener_Correos");
                    command.Parameters.AddWithValue("@DLType", "Notificacion_Fallas");
                    
                    conn.Open();

                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(reader.GetOrdinal("disM_Detalles")))
                            {
                                recipients = reader.GetString(reader.GetOrdinal("disM_Detalles"));
                            }
                            else
                            {
                                recipients = "";
                            }
                        }
                    }
                }
            }

            return recipients;
        }

        public static void SendEmailWithCsvData(string recipients, string subject, List<ResultRecord> results, CSVRecord header)
        {

            using (var smtpClient = new SmtpClient("mail.apise.com.mx"))
            {
                smtpClient.Port = 587;
                smtpClient.Credentials = new NetworkCredential("jrios@apise.com.mx", "Rwe3xws1324");
                smtpClient.EnableSsl = false;

                var mailMessage = new MailMessage();
                mailMessage.From = new MailAddress("jrios@apise.com.mx");

                mailMessage.To.Add(recipients);

                //foreach (var recipient in recipients)
                //{
                //    mailMessage.To.Add(recipient);
                //}

                mailMessage.Subject = subject;
                mailMessage.Body = $"<h2>Resultados de Pruebas de Producto</h2>{ConvertToHtmlTable(results,header)}";
                mailMessage.IsBodyHtml = true; // O true si quieres enviar el correo en HTML

                smtpClient.Send(mailMessage);
            }
        }

        public static string ConvertToHtmlTable(List<ResultRecord> results, CSVRecord header)
        {
            var sb = new StringBuilder();

            sb.Append("<table border='1'>");
            sb.Append("<tr>");
            sb.Append("<th>Cal. Date</th>");
            sb.Append("<th>Equipo</th>");
            sb.Append("<th>Inspector</th>");
            sb.Append("<th>Lado</th>");
            sb.Append("<th>Lote</th>");
            sb.Append("<th># de Parte</th>");
            sb.Append("<th>Sello Pel</th>");
            sb.Append("<th>Result</th>");
            sb.Append("<th>Max Load</th>");
            sb.Append("<th>Estado</th>");
            sb.Append("</tr>");

            int totalTests = results.Count;
            int successfulTests = results.Count(r => r.Estado == "Paso");
            int failedTests = results.Count(r => r.Estado == "Falla");

            foreach (var r in results)
            {
                sb.Append("<tr>");
                sb.AppendFormat("<td>{0}</td>", header.CalDate);
                sb.AppendFormat("<td>{0}</td>", header.Equipo);
                sb.AppendFormat("<td>{0}</td>", header.Inspector);
                sb.AppendFormat("<td>{0}</td>", header.Lado);
                sb.AppendFormat("<td>{0}</td>", header.Lote);
                sb.AppendFormat("<td>{0}</td>", header.NumeroDeParte);
                sb.AppendFormat("<td>{0}</td>", header.SelloPel);
                sb.AppendFormat("<td>{0}</td>", r.Result);
                sb.AppendFormat("<td>{0}</td>", r.MaximumLoad);

                // Verificar el valor de r.Estado y agregar un estilo si es "Falla"
                if (r.Estado == "Falla")
                {
                    sb.AppendFormat("<td style='background-color: red; color: white; font-weight: bold;'>{0}</td>", r.Estado);
                }
                else
                {
                    sb.AppendFormat("<td style='background-color: green; color: white; font-weight: bold;'>{0}</td>", r.Estado);
                }
                sb.Append("</tr>");
            }
            sb.Append("</table>");

            // Agregar el resumen al final
            sb.Append("<p><strong>Resumen de Pruebas:</strong></p>");
            sb.Append("<ul>");
            sb.AppendFormat("<li>Cantidad de Pruebas: {0}</li>", totalTests);
            sb.AppendFormat("<li>Cantidad de Pruebas Exitosas: {0}</li>", successfulTests);
            sb.AppendFormat("<li>Cantidad de Fallas: {0}</li>", failedTests);
            sb.Append("</ul>");

            return sb.ToString();
        }
    }

}
