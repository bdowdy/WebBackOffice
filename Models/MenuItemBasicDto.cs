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
    public decimal LiquorTaxAmount { get; set; }

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
    public bool MinMaxChoices { get; set; }
    public bool OpenPrice { get; set; }
    public bool QuickPick { get; set; }
    public bool QuickButton { get; set; }
    public bool OnLine { get; set; }

    // Timed Pricing
    public bool TimedPrice { get; set; }

    public string TimeEvent1 { get; set; } = "";
    public string TimeEvent2 { get; set; } = "";
    public string TimeEvent3 { get; set; } = "";
    public string TimeEvent4 { get; set; } = "";
    public string TimeEvent5 { get; set; } = "";

    public decimal TimePrice1 { get; set; }
    public decimal TimePrice2 { get; set; }
    public decimal TimePrice3 { get; set; }
    public decimal TimePrice4 { get; set; }
    public decimal TimePrice5 { get; set; }

    // Modifier Assignments
    public string Modifier1 { get; set; } = "";
    public string Modifier2 { get; set; } = "";
    public string Modifier3 { get; set; } = "";
    public string Modifier4 { get; set; } = "";
    public string Modifier5 { get; set; } = "";
    public string Modifier6 { get; set; } = "";

    public int ModifierOption1 { get; set; }
    public int ModifierOption2 { get; set; }
    public int ModifierOption3 { get; set; }
    public int ModifierOption4 { get; set; }
    public int ModifierOption5 { get; set; }
    public int ModifierOption6 { get; set; }

    // These come from [Menu Mods]
    public int ModifierMin1 { get; set; }
    public int ModifierMin2 { get; set; }
    public int ModifierMin3 { get; set; }
    public int ModifierMin4 { get; set; }
    public int ModifierMin5 { get; set; }
    public int ModifierMin6 { get; set; }

    public int ModifierMax1 { get; set; }
    public int ModifierMax2 { get; set; }
    public int ModifierMax3 { get; set; }
    public int ModifierMax4 { get; set; }
    public int ModifierMax5 { get; set; }
    public int ModifierMax6 { get; set; }
    public int ModifierGroupStart { get; set; } = 1;

    // Shift Pricing
    public bool PriceLevels { get; set; }
    public bool PriceLock { get; set; }

    public decimal Price1 { get; set; }
    public decimal Price2 { get; set; }
    public decimal Price3 { get; set; }
    public decimal Price4 { get; set; }
    public decimal Price5 { get; set; }

    public string ShiftEvent1 { get; set; } = "";
    public string ShiftEvent2 { get; set; } = "";
    public string ShiftEvent3 { get; set; } = "";
    public string ShiftEvent4 { get; set; } = "";
    public string ShiftEvent5 { get; set; } = "";

    public decimal ShiftPrice1 { get; set; }
    public decimal ShiftPrice2 { get; set; }
    public decimal ShiftPrice3 { get; set; }
    public decimal ShiftPrice4 { get; set; }
    public decimal ShiftPrice5 { get; set; }

    public string ShiftMod1 { get; set; } = "";
    public string ShiftMod2 { get; set; } = "";
    public string ShiftMod3 { get; set; } = "";
    public string ShiftMod4 { get; set; } = "";
    public string ShiftMod5 { get; set; } = "";

    public int EModShift1 { get; set; }
    public int EModShift2 { get; set; }
    public int EModShift3 { get; set; }
    public int EModShift4 { get; set; }
    public int EModShift5 { get; set; }

    public int PModShift1 { get; set; }
    public int PModShift2 { get; set; }
    public int PModShift3 { get; set; }
    public int PModShift4 { get; set; }
    public int PModShift5 { get; set; }

    public int ShiftPLU1 { get; set; }
    public int ShiftPLU2 { get; set; }
    public int ShiftPLU3 { get; set; }
    public int ShiftPLU4 { get; set; }
    public int ShiftPLU5 { get; set; }
}
