namespace MSaleWebServer.Models
{
    public class DiscountEditModel
    {
        public string DiscountName { get; set; } = "";
        public string DiscountType { get; set; } = "";
        public decimal DiscountAmount { get; set; }
        public string AccessLevel { get; set; } = "";

        public string CategoryLink { get; set; } = "";
        public string MenuItemLink { get; set; } = "";

        public decimal MaxDiscountAmount { get; set; }

        public bool Taxable { get; set; } = false;
        public string DepartmentLink { get; set; } = "";

        public string StartTime { get; set; } = "";
        public string EndTime { get; set; } = "";
        public string Expires { get; set; } = "";

       
        public string DiscountCategory { get; set; } = "";
        public int Tax1 { get; set; } = 0;
        public int Tax2 { get; set; } = 0;
        public bool AllowOnlyOne { get; set; }
        public bool AfterPrint { get; set; } // maps to exclude modifiers
        public bool AutoMatch { get; set; }
        public bool AutoDeduct { get; set; }
        public bool Tracking { get; set; }
        public bool EventOnly { get; set; } // maps to [Search Number]

    }
}
