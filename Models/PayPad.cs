using System;

namespace API_INDER_INFORMES.Models
{
    public class PayPad
    {
        public int ID { get; set; }
        public string USERNAME { get; set; }
        public string PWD { get; set; }
        public string DESCRIPTION { get; set; }
        public double? LONGITUDE { get; set; }
        public double? LATITUDE { get; set; }
        public string STATUS { get; set; }
        public int? ID_CURRENCY { get; set; }
        public int? ID_OFFICE { get; set; }
        public int? ID_USER_CREATED { get; set; }
        public DateTime DATE_CREATED { get; set; }
        public int? ID_USER_UPDATED { get; set; }
        public DateTime? DATE_UPDATED { get; set; }
    }
}
