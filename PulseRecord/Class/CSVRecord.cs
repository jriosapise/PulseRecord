using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PulseRecord.Class
{
    public class CSVRecord
    {
        public string CalDate { get; set; }
        public string Equipo { get; set; }
        public string Inspector { get; set; }
        public string Lado { get; set; }
        public string Lote { get; set; }
        public string NumeroDeParte { get; set; }
        public string SelloPel { get; set; }

        // Lista para almacenar los registros de resultados
        public List<ResultRecord> Results { get; set; }

        public CSVRecord()
        {
            Results = new List<ResultRecord>();
        }
    }
    public class ResultRecord
    {
        public int Result { get; set; }
        public double MaximumLoad { get; set; }
        public string Estado { get; set; }
    }

    public class MinMaxValores
    {
        public double Minimo { get; set; }
        public double Maximo { get; set; }
    }

}
