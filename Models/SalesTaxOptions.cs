    namespace MSaleWebServer.Models
    {
        public class SalesTaxOption
        {
            public string TaxName { get; set; } = "";
            public decimal TaxAmount { get; set; }
            public int TaxTable { get; set; }
        }
    }
