using System;

namespace API_INDER_INFORMES.Models
{
    public class StateTransaction
    {
        public int ID { get; set; }
        public string STATE { get; set; }
        public int? ID_USER_CREATED { get; set; }
        public DateTime DATE_CREATED { get; set; }
        public int? ID_USER_UPDATED { get; set; }
        public DateTime? DATE_UPDATED { get; set; }
    }
}
