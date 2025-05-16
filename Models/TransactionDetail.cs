using System;

namespace API_INDER_INFORMES.Models
{
    public class TransactionDetail
    {
        public int ID { get; set; }
        public int ID_TRANSACTION { get; set; }
        public int ID_CURRENCY_DENOMINATION { get; set; }
        public int ID_TYPE_OPERATION { get; set; }
        public DateTime DATE_CREATED { get; set; }
        public DateTime? DATE_UPDATED { get; set; }
    }
}
