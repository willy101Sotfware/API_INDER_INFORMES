using System;

namespace API_INDER_INFORMES.Models
{
    public class Transaction
    {
        public int ID { get; set; }
        public string? DOCUMENT { get; set; }
        public string? REFERENCE { get; set; }
        public string? PRODUCT { get; set; }
        public double? TOTAL_AMOUNT { get; set; }
        public double? REAL_AMOUNT { get; set; }
        public double? INCOME_AMOUNT { get; set; }
        public double? RETURN_AMOUNT { get; set; }
        public string? DESCRIPTION { get; set; }
        public int? ID_STATE_TRANSACTION { get; set; }
        public int? ID_TYPE_TRANSACTION { get; set; }
        public int? ID_TYPE_PAYMENT { get; set; }
        public int? ID_PAYPAD { get; set; }
        public DateTime DATE_CREATED { get; set; }
        public DateTime? DATE_UPDATED { get; set; }
    }
}
