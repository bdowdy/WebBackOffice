namespace MSaleWebServer.Models
{
    public class ModifierDetailDto
    {
        public bool IsNewItem { get; set; }
        public string ModifierName { get; set; }
        public string ModifierCategory { get; set; }

        public string OriginalModifierName { get; set; }
        public string OriginalModifierCategory { get; set; }

        public decimal ModifierCost { get; set; }
        public decimal ItemCost { get; set; }

        public int SalesTax { get; set; }
        public int SpecialTax { get; set; }
        public int LiquorTax { get; set; }

        public string RemotePrinter1 { get; set; }
        public string RemotePrinter2 { get; set; }
        public string RemotePrinter3 { get; set; }
        public string RemotePrinter4 { get; set; }
        public string RemotePrinter5 { get; set; }
        public string RemotePrinter6 { get; set; }

        public string PriceLevel { get; set; }
        public string AltModifier { get; set; }
        public int SelectionQuantity { get; set; }
        public string Alias { get; set; }
        public string ComesWith { get; set; }
        public string PLU { get; set; }
    }
}
