namespace MSaleWebServer.Models
{
    public class TaxSetupSaveRequest
    {
        public List<SalesTaxDto> SalesTaxes { get; set; } = new();
        public List<SpecialTaxDto> SpecialTaxes { get; set; } = new();
    }

    public class SalesTaxDto
    {
        public int TaxTable { get; set; }
        public string TaxName { get; set; } = "";
        public decimal TaxAmount { get; set; }
    }

    public class SpecialTaxDto
    {
        public int TaxTable { get; set; }
        public string TaxName { get; set; } = "";
        public decimal SpecialTax { get; set; }
        public decimal MinTax { get; set; }
        public decimal MaxTax { get; set; }
        public string Exception { get; set; } = "";
    }
}
