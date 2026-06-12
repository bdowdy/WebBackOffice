namespace MSaleWebServer.Models
{
    public class MenuItemBasicDto
    {
        public bool IsNewItem { get; set; }

        public string MenuItemName { get; set; } = "";
        public string OriginalMenuItemName { get; set; } = "";

        public string DepartmentName { get; set; } = "";
        public string ReportCategory { get; set; } = "";

        public decimal RegularPrice { get; set; }
        public decimal InventoryAmount { get; set; }
        public decimal ProductCost { get; set; }
        public decimal UpCharge { get; set; }

        public string BarCode { get; set; } = "";
        public string ItemInfo { get; set; } = "";

        public int SalesTaxTable { get; set; }
        public int SpecialTaxTable { get; set; }
        public int LiquorTaxAmount { get; set; }

        public string Remote1 { get; set; } = "";
        public string Remote2 { get; set; } = "";
        public string Remote3 { get; set; } = "";
        public string Remote4 { get; set; } = "";
        public string Remote5 { get; set; } = "";
        public string Remote6 { get; set; } = "";

        public bool Avalibility { get; set; }
        public bool Taxable { get; set; }
        public bool UseModifiers { get; set; }
        public bool ModifiersAreOptional { get; set; }
        public bool OpenPrice { get; set; }
        public bool QuickPick { get; set; }
        public bool QuickButton { get; set; }
        public bool OnLine { get; set; }
    }
}