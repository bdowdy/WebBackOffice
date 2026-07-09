using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MSaleWebServer.Models;

namespace MSaleWebServer.Controllers
{
    public class MenuItemController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly Services.TerminalUpdateService _terminals;

        public MenuItemController(IConfiguration configuration, Services.TerminalUpdateService terminals)
        {
            _configuration = configuration;
            _terminals = terminals;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View("~/Views/MSaleAdmin/MenuItem.cshtml");
        }

        [HttpGet]
        public IActionResult GetMenuItemNames()
        {
            var list = new List<string>();

            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();

            using var cmd = new SqlCommand(@"
                SELECT [Menu Item Name]
                FROM [Menu Items]
                WHERE [Menu Item Name] IS NOT NULL
                  AND LTRIM(RTRIM([Menu Item Name])) <> ''
                ORDER BY [Menu Item Name]", cn);

            using var dr = cmd.ExecuteReader();

            while (dr.Read())
                list.Add(dr["Menu Item Name"].ToString() ?? "");

            return Json(list);
        }

        [HttpGet]
        public IActionResult GetMenuItem(string menuItemName)
        {
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();

            using var cmd = new SqlCommand(@"
        SELECT TOP 1
            [Menu Item Name],
            ISNULL([Department Name], '') AS [Department Name],
            ISNULL([Report Category], '') AS [Report Category],
            ISNULL([Regular Price], 0) AS [Regular Price],
            ISNULL([Inventory Amount], 0) AS [Inventory Amount],
            ISNULL([Product Cost], 0) AS [Product Cost],
            ISNULL([Up Charge], 0) AS [Up Charge],
            ISNULL([Bar Code], '') AS [Bar Code],
            ISNULL([Item Info], '') AS [Item Info],

            ISNULL([Sales Tax Table], 0) AS [Sales Tax Table],
            ISNULL([Special Tax Table], 0) AS [Special Tax Table],
            ISNULL([Liquor Tax Amount], 0) AS [Liquor Tax Amount],

            ISNULL([Remote 1], '') AS [Remote 1],
            ISNULL([Remote 2], '') AS [Remote 2],
            ISNULL([Remote 3], '') AS [Remote 3],
            ISNULL([Remote 4], '') AS [Remote 4],
            ISNULL([Remote 5], '') AS [Remote 5],
            ISNULL([Remote 6], '') AS [Remote 6],

            ISNULL([Avalibility], 0) AS [Avalibility],
            ISNULL([Taxable], 0) AS [Taxable],
            ISNULL([Use Modifiers], 0) AS [Use Modifiers],
            ISNULL([Modifiers are Optional], 0) AS [Modifiers are Optional],
            ISNULL([Open Price], 0) AS [Open Price],
            ISNULL([Quick Pick], 0) AS [Quick Pick],
            ISNULL([Quick Button], 0) AS [Quick Button],
            ISNULL([On Line], 0) AS [On Line],

            ISNULL([Timed Price], 0) AS [Timed Price],
            ISNULL([Time Event 1], '') AS [Time Event 1],
            ISNULL([Time Event 2], '') AS [Time Event 2],
            ISNULL([Time Event 3], '') AS [Time Event 3],
            ISNULL([Time Event 4], '') AS [Time Event 4],
            ISNULL([Time Event 5], '') AS [Time Event 5],
            ISNULL([Time Price 1], 0) AS [Time Price 1],
            ISNULL([Time Price 2], 0) AS [Time Price 2],
            ISNULL([Time Price 3], 0) AS [Time Price 3],
            ISNULL([Time Price 4], 0) AS [Time Price 4],
            ISNULL([Time Price 5], 0) AS [Time Price 5],

            ISNULL([Min Max Choices], 0) AS [Min Max Choices],
            ISNULL([Modifier 1], '') AS [Modifier 1],
            ISNULL([Modifier 2], '') AS [Modifier 2],
            ISNULL([Modifier 3], '') AS [Modifier 3],
            ISNULL([Modifier 4], '') AS [Modifier 4],
            ISNULL([Modifier 5], '') AS [Modifier 5],
            ISNULL([Modifier 6], '') AS [Modifier 6],
            ISNULL([Modifier Option 1], 0) AS [Modifier Option 1],
            ISNULL([Modifier Option 2], 0) AS [Modifier Option 2],
            ISNULL([Modifier Option 3], 0) AS [Modifier Option 3],
            ISNULL([Modifier Option 4], 0) AS [Modifier Option 4],
            ISNULL([Modifier Option 5], 0) AS [Modifier Option 5],
            ISNULL([Modifier Option 6], 0) AS [Modifier Option 6],

            ISNULL([Price Levels], 0) AS [Price Levels],
            ISNULL([price Lock], 0) AS [price Lock],
            ISNULL([Price 1], 0) AS [Price 1],
            ISNULL([Price 2], 0) AS [Price 2],
            ISNULL([Price 3], 0) AS [Price 3],
            ISNULL([Price 4], 0) AS [Price 4],
            ISNULL([Price 5], 0) AS [Price 5],

            ISNULL([Shift Event 1], '') AS [Shift Event 1],
            ISNULL([Shift Event 2], '') AS [Shift Event 2],
            ISNULL([Shift Event 3], '') AS [Shift Event 3],
            ISNULL([Shift Event 4], '') AS [Shift Event 4],
            ISNULL([Shift Event 5], '') AS [Shift Event 5],
            ISNULL([Shift Price 1], 0) AS [Shift Price 1],
            ISNULL([Shift Price 2], 0) AS [Shift Price 2],
            ISNULL([Shift Price 3], 0) AS [Shift Price 3],
            ISNULL([Shift Price 4], 0) AS [Shift Price 4],
            ISNULL([Shift Price 5], 0) AS [Shift Price 5],
            ISNULL([Shift Mod 1], '') AS [Shift Mod 1],
            ISNULL([Shift Mod 2], '') AS [Shift Mod 2],
            ISNULL([Shift Mod 3], '') AS [Shift Mod 3],
            ISNULL([Shift Mod 4], '') AS [Shift Mod 4],
            ISNULL([Shift Mod 5], '') AS [Shift Mod 5],
            ISNULL([EModShift 1], 0) AS [EModShift 1],
            ISNULL([EModShift 2], 0) AS [EModShift 2],
            ISNULL([EModShift 3], 0) AS [EModShift 3],
            ISNULL([EModShift 4], 0) AS [EModShift 4],
            ISNULL([EModShift 5], 0) AS [EModShift 5],
            ISNULL([PModShift 1], 0) AS [PModShift 1],
            ISNULL([PModShift 2], 0) AS [PModShift 2],
            ISNULL([PModShift 3], 0) AS [PModShift 3],
            ISNULL([PModShift 4], 0) AS [PModShift 4],
            ISNULL([PModShift 5], 0) AS [PModShift 5],
            ISNULL([Shift 1 PLU], 0) AS [Shift 1 PLU],
            ISNULL([Shift 2 PLU], 0) AS [Shift 2 PLU],
            ISNULL([Shift 3 PLU], 0) AS [Shift 3 PLU],
            ISNULL([Shift 4 PLU], 0) AS [Shift 4 PLU],
            ISNULL([Shift 5 PLU], 0) AS [Shift 5 PLU]

        FROM [Menu Items]
        WHERE [Menu Item Name] = @MenuItemName", cn);

            cmd.Parameters.AddWithValue("@MenuItemName", menuItemName ?? "");

            MenuItemBasicDto item;

            using (var dr = cmd.ExecuteReader())
            {
                if (!dr.Read())
                    return Json(new { success = false, message = "Menu item not found." });

                item = new MenuItemBasicDto
                {
                    IsNewItem = false,

                    MenuItemName = dr["Menu Item Name"].ToString() ?? "",
                    OriginalMenuItemName = dr["Menu Item Name"].ToString() ?? "",

                    DepartmentName = dr["Department Name"].ToString() ?? "",
                    ReportCategory = dr["Report Category"].ToString() ?? "",

                    RegularPrice = Convert.ToDecimal(dr["Regular Price"]),
                    InventoryAmount = Convert.ToDecimal(dr["Inventory Amount"]),
                    ProductCost = Convert.ToDecimal(dr["Product Cost"]),
                    UpCharge = Convert.ToDecimal(dr["Up Charge"]),

                    BarCode = dr["Bar Code"].ToString() ?? "",
                    ItemInfo = dr["Item Info"].ToString() ?? "",

                    SalesTaxTable = Convert.ToInt32(dr["Sales Tax Table"]),
                    SpecialTaxTable = Convert.ToInt32(dr["Special Tax Table"]),
                    LiquorTaxAmount = Convert.ToDecimal(dr["Liquor Tax Amount"]),

                    Remote1 = dr["Remote 1"].ToString() ?? "",
                    Remote2 = dr["Remote 2"].ToString() ?? "",
                    Remote3 = dr["Remote 3"].ToString() ?? "",
                    Remote4 = dr["Remote 4"].ToString() ?? "",
                    Remote5 = dr["Remote 5"].ToString() ?? "",
                    Remote6 = dr["Remote 6"].ToString() ?? "",

                    Avalibility = Convert.ToBoolean(dr["Avalibility"]),
                    Taxable = Convert.ToBoolean(dr["Taxable"]),
                    UseModifiers = Convert.ToBoolean(dr["Use Modifiers"]),
                    ModifiersAreOptional = Convert.ToBoolean(dr["Modifiers are Optional"]),
                    OpenPrice = Convert.ToBoolean(dr["Open Price"]),
                    QuickPick = Convert.ToBoolean(dr["Quick Pick"]),
                    QuickButton = Convert.ToBoolean(dr["Quick Button"]),
                    OnLine = Convert.ToBoolean(dr["On Line"]),

                    TimedPrice = Convert.ToBoolean(dr["Timed Price"]),
                    TimeEvent1 = dr["Time Event 1"].ToString() ?? "",
                    TimeEvent2 = dr["Time Event 2"].ToString() ?? "",
                    TimeEvent3 = dr["Time Event 3"].ToString() ?? "",
                    TimeEvent4 = dr["Time Event 4"].ToString() ?? "",
                    TimeEvent5 = dr["Time Event 5"].ToString() ?? "",
                    TimePrice1 = Convert.ToDecimal(dr["Time Price 1"]),
                    TimePrice2 = Convert.ToDecimal(dr["Time Price 2"]),
                    TimePrice3 = Convert.ToDecimal(dr["Time Price 3"]),
                    TimePrice4 = Convert.ToDecimal(dr["Time Price 4"]),
                    TimePrice5 = Convert.ToDecimal(dr["Time Price 5"]),

                    MinMaxChoices = Convert.ToBoolean(dr["Min Max Choices"]),

                    Modifier1 = dr["Modifier 1"].ToString() ?? "",
                    Modifier2 = dr["Modifier 2"].ToString() ?? "",
                    Modifier3 = dr["Modifier 3"].ToString() ?? "",
                    Modifier4 = dr["Modifier 4"].ToString() ?? "",
                    Modifier5 = dr["Modifier 5"].ToString() ?? "",
                    Modifier6 = dr["Modifier 6"].ToString() ?? "",

                    ModifierOption1 = Convert.ToInt32(dr["Modifier Option 1"]),
                    ModifierOption2 = Convert.ToInt32(dr["Modifier Option 2"]),
                    ModifierOption3 = Convert.ToInt32(dr["Modifier Option 3"]),
                    ModifierOption4 = Convert.ToInt32(dr["Modifier Option 4"]),
                    ModifierOption5 = Convert.ToInt32(dr["Modifier Option 5"]),
                    ModifierOption6 = Convert.ToInt32(dr["Modifier Option 6"]),

                    PriceLevels = Convert.ToBoolean(dr["Price Levels"]),
                    PriceLock = Convert.ToBoolean(dr["price Lock"]),
                    Price1 = Convert.ToDecimal(dr["Price 1"]),
                    Price2 = Convert.ToDecimal(dr["Price 2"]),
                    Price3 = Convert.ToDecimal(dr["Price 3"]),
                    Price4 = Convert.ToDecimal(dr["Price 4"]),
                    Price5 = Convert.ToDecimal(dr["Price 5"]),

                    ShiftEvent1 = dr["Shift Event 1"].ToString() ?? "",
                    ShiftEvent2 = dr["Shift Event 2"].ToString() ?? "",
                    ShiftEvent3 = dr["Shift Event 3"].ToString() ?? "",
                    ShiftEvent4 = dr["Shift Event 4"].ToString() ?? "",
                    ShiftEvent5 = dr["Shift Event 5"].ToString() ?? "",
                    ShiftPrice1 = Convert.ToDecimal(dr["Shift Price 1"]),
                    ShiftPrice2 = Convert.ToDecimal(dr["Shift Price 2"]),
                    ShiftPrice3 = Convert.ToDecimal(dr["Shift Price 3"]),
                    ShiftPrice4 = Convert.ToDecimal(dr["Shift Price 4"]),
                    ShiftPrice5 = Convert.ToDecimal(dr["Shift Price 5"]),
                    ShiftMod1 = dr["Shift Mod 1"].ToString() ?? "",
                    ShiftMod2 = dr["Shift Mod 2"].ToString() ?? "",
                    ShiftMod3 = dr["Shift Mod 3"].ToString() ?? "",
                    ShiftMod4 = dr["Shift Mod 4"].ToString() ?? "",
                    ShiftMod5 = dr["Shift Mod 5"].ToString() ?? "",
                    EModShift1 = Convert.ToInt32(dr["EModShift 1"]),
                    EModShift2 = Convert.ToInt32(dr["EModShift 2"]),
                    EModShift3 = Convert.ToInt32(dr["EModShift 3"]),
                    EModShift4 = Convert.ToInt32(dr["EModShift 4"]),
                    EModShift5 = Convert.ToInt32(dr["EModShift 5"]),
                    PModShift1 = Convert.ToInt32(dr["PModShift 1"]),
                    PModShift2 = Convert.ToInt32(dr["PModShift 2"]),
                    PModShift3 = Convert.ToInt32(dr["PModShift 3"]),
                    PModShift4 = Convert.ToInt32(dr["PModShift 4"]),
                    PModShift5 = Convert.ToInt32(dr["PModShift 5"]),
                    ShiftPLU1 = Convert.ToInt32(dr["Shift 1 PLU"]),
                    ShiftPLU2 = Convert.ToInt32(dr["Shift 2 PLU"]),
                    ShiftPLU3 = Convert.ToInt32(dr["Shift 3 PLU"]),
                    ShiftPLU4 = Convert.ToInt32(dr["Shift 4 PLU"]),
                    ShiftPLU5 = Convert.ToInt32(dr["Shift 5 PLU"])
                };
            }

            using var modCmd = new SqlCommand(@"
        SELECT 
            ISNULL([Modifier Name], '') AS [Modifier Name],
            ISNULL([Min], 0) AS [Min],
            ISNULL([Max], 0) AS [Max]
        FROM [Menu Mods]
        WHERE [Menu Item Name] = @MenuItemName", cn);

            modCmd.Parameters.AddWithValue("@MenuItemName", item.MenuItemName);

            using var modDr = modCmd.ExecuteReader();

            while (modDr.Read())
            {
                string modName = modDr["Modifier Name"].ToString() ?? "";
                int min = Convert.ToInt32(modDr["Min"]);
                int max = Convert.ToInt32(modDr["Max"]);

                if (modName == item.Modifier1) { item.ModifierMin1 = min; item.ModifierMax1 = max; }
                else if (modName == item.Modifier2) { item.ModifierMin2 = min; item.ModifierMax2 = max; }
                else if (modName == item.Modifier3) { item.ModifierMin3 = min; item.ModifierMax3 = max; }
                else if (modName == item.Modifier4) { item.ModifierMin4 = min; item.ModifierMax4 = max; }
                else if (modName == item.Modifier5) { item.ModifierMin5 = min; item.ModifierMax5 = max; }
                else if (modName == item.Modifier6) { item.ModifierMin6 = min; item.ModifierMax6 = max; }
            }

            return Json(new { success = true, data = item });
        }

        [HttpPost]
        public IActionResult SaveMenuItem([FromBody] MenuItemBasicDto model)
        {
            try
            {
                if (model == null)
                    return Json(new { success = false, message = "No data received." });

                if (string.IsNullOrWhiteSpace(model.MenuItemName))
                    return Json(new { success = false, message = "Menu item name is required." });

                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();

                using var tran = cn.BeginTransaction();

                try
                {
                    int rows;

                    if (model.IsNewItem)
                    {
                        using var check = new SqlCommand(@"
                    SELECT COUNT(*)
                    FROM [Menu Items]
                    WHERE [Menu Item Name] = @MenuItemName", cn, tran);

                        check.Parameters.AddWithValue("@MenuItemName", model.MenuItemName.Trim());

                        int exists = Convert.ToInt32(check.ExecuteScalar());

                        if (exists > 0)
                        {
                            tran.Rollback();
                            return Json(new { success = false, message = "That menu item already exists." });
                        }

                        using var insert = new SqlCommand(@"
                    INSERT INTO [Menu Items]
                    (
                        [Menu Item Name],
                        [Department Name],
                        [Report Category],
                        [Regular Price],
                        [Inventory Amount],
                        [Product Cost],
                        [Up Charge],
                        [Bar Code],
                        [Item Info],

                        [Sales Tax Table],
                        [Special Tax Table],
                        [Liquor Tax Amount],

                        [Remote 1],
                        [Remote 2],
                        [Remote 3],
                        [Remote 4],
                        [Remote 5],
                        [Remote 6],

                        [Avalibility],
                        [Taxable],
                        [Use Modifiers],
                        [Modifiers are Optional],
                        [Min Max Choices],
                        [Open Price],
                        [Quick Pick],
                        [Quick Button],
                        [On Line],
                        [Timed Price],

                        [Time Event 1],
                        [Time Event 2],
                        [Time Event 3],
                        [Time Event 4],
                        [Time Event 5],
                        [Time Price 1],
                        [Time Price 2],
                        [Time Price 3],
                        [Time Price 4],
                        [Time Price 5],

                        [Modifier 1],
                        [Modifier 2],
                        [Modifier 3],
                        [Modifier 4],
                        [Modifier 5],
                        [Modifier 6],
                        [Modifier Option 1],
                        [Modifier Option 2],
                        [Modifier Option 3],
                        [Modifier Option 4],
                        [Modifier Option 5],
                        [Modifier Option 6],

                        [Price Levels],
                        [price Lock],
                        [Price 1],
                        [Price 2],
                        [Price 3],
                        [Price 4],
                        [Price 5],

                        [Shift Event 1],
                        [Shift Event 2],
                        [Shift Event 3],
                        [Shift Event 4],
                        [Shift Event 5],
                        [Shift Price 1],
                        [Shift Price 2],
                        [Shift Price 3],
                        [Shift Price 4],
                        [Shift Price 5],
                        [Shift Mod 1],
                        [Shift Mod 2],
                        [Shift Mod 3],
                        [Shift Mod 4],
                        [Shift Mod 5],
                        [EModShift 1],
                        [EModShift 2],
                        [EModShift 3],
                        [EModShift 4],
                        [EModShift 5],
                        [PModShift 1],
                        [PModShift 2],
                        [PModShift 3],
                        [PModShift 4],
                        [PModShift 5],
                        [Shift 1 PLU],
                        [Shift 2 PLU],
                        [Shift 3 PLU],
                        [Shift 4 PLU],
                        [Shift 5 PLU],

                        [PLU],
                        [Record Number]
                    )
                    VALUES
                    (
                        @MenuItemName,
                        @DepartmentName,
                        @ReportCategory,
                        @RegularPrice,
                        @InventoryAmount,
                        @ProductCost,
                        @UpCharge,
                        @BarCode,
                        @ItemInfo,

                        @SalesTaxTable,
                        @SpecialTaxTable,
                        @LiquorTaxAmount,

                        @Remote1,
                        @Remote2,
                        @Remote3,
                        @Remote4,
                        @Remote5,
                        @Remote6,

                        @Avalibility,
                        @Taxable,
                        @UseModifiers,
                        @ModifiersAreOptional,
                        @MinMaxChoices,
                        @OpenPrice,
                        @QuickPick,
                        @QuickButton,
                        @OnLine,
                        @TimedPrice,

                        @TimeEvent1,
                        @TimeEvent2,
                        @TimeEvent3,
                        @TimeEvent4,
                        @TimeEvent5,
                        @TimePrice1,
                        @TimePrice2,
                        @TimePrice3,
                        @TimePrice4,
                        @TimePrice5,

                        @Modifier1,
                        @Modifier2,
                        @Modifier3,
                        @Modifier4,
                        @Modifier5,
                        @Modifier6,
                        @ModifierOption1,
                        @ModifierOption2,
                        @ModifierOption3,
                        @ModifierOption4,
                        @ModifierOption5,
                        @ModifierOption6,

                        @PriceLevels,
                        @PriceLock,
                        @Price1,
                        @Price2,
                        @Price3,
                        @Price4,
                        @Price5,

                        @ShiftEvent1,
                        @ShiftEvent2,
                        @ShiftEvent3,
                        @ShiftEvent4,
                        @ShiftEvent5,
                        @ShiftPrice1,
                        @ShiftPrice2,
                        @ShiftPrice3,
                        @ShiftPrice4,
                        @ShiftPrice5,
                        @ShiftMod1,
                        @ShiftMod2,
                        @ShiftMod3,
                        @ShiftMod4,
                        @ShiftMod5,
                        @EModShift1,
                        @EModShift2,
                        @EModShift3,
                        @EModShift4,
                        @EModShift5,
                        @PModShift1,
                        @PModShift2,
                        @PModShift3,
                        @PModShift4,
                        @PModShift5,
                        @ShiftPLU1,
                        @ShiftPLU2,
                        @ShiftPLU3,
                        @ShiftPLU4,
                        @ShiftPLU5,

                        -1,
                        ISNULL((SELECT MAX([Record Number]) + 1 FROM [Menu Items]), 1)
                    )", cn, tran);

                        AddParameters(insert, model);
                        rows = insert.ExecuteNonQuery();
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(model.OriginalMenuItemName))
                        {
                            tran.Rollback();
                            return Json(new { success = false, message = "Original menu item name is missing." });
                        }

                        using var update = new SqlCommand(@"
                    UPDATE [Menu Items]
                    SET
                        [Menu Item Name] = @MenuItemName,
                        [Department Name] = @DepartmentName,
                        [Report Category] = @ReportCategory,
                        [Regular Price] = @RegularPrice,
                        [Inventory Amount] = @InventoryAmount,
                        [Product Cost] = @ProductCost,
                        [Up Charge] = @UpCharge,
                        [Bar Code] = @BarCode,
                        [Item Info] = @ItemInfo,

                        [Sales Tax Table] = @SalesTaxTable,
                        [Special Tax Table] = @SpecialTaxTable,
                        [Liquor Tax Amount] = @LiquorTaxAmount,

                        [Remote 1] = @Remote1,
                        [Remote 2] = @Remote2,
                        [Remote 3] = @Remote3,
                        [Remote 4] = @Remote4,
                        [Remote 5] = @Remote5,
                        [Remote 6] = @Remote6,

                        [Avalibility] = @Avalibility,
                        [Taxable] = @Taxable,
                        [Use Modifiers] = @UseModifiers,
                        [Modifiers are Optional] = @ModifiersAreOptional,
                        [Min Max Choices] = @MinMaxChoices,
                        [Open Price] = @OpenPrice,
                        [Quick Pick] = @QuickPick,
                        [Quick Button] = @QuickButton,
                        [On Line] = @OnLine,
                        [Timed Price] = @TimedPrice,

                        [Time Event 1] = @TimeEvent1,
                        [Time Event 2] = @TimeEvent2,
                        [Time Event 3] = @TimeEvent3,
                        [Time Event 4] = @TimeEvent4,
                        [Time Event 5] = @TimeEvent5,
                        [Time Price 1] = @TimePrice1,
                        [Time Price 2] = @TimePrice2,
                        [Time Price 3] = @TimePrice3,
                        [Time Price 4] = @TimePrice4,
                        [Time Price 5] = @TimePrice5,

                        [Modifier 1] = @Modifier1,
                        [Modifier 2] = @Modifier2,
                        [Modifier 3] = @Modifier3,
                        [Modifier 4] = @Modifier4,
                        [Modifier 5] = @Modifier5,
                        [Modifier 6] = @Modifier6,
                        [Modifier Option 1] = @ModifierOption1,
                        [Modifier Option 2] = @ModifierOption2,
                        [Modifier Option 3] = @ModifierOption3,
                        [Modifier Option 4] = @ModifierOption4,
                        [Modifier Option 5] = @ModifierOption5,
                        [Modifier Option 6] = @ModifierOption6,

                        [Price Levels] = @PriceLevels,
                        [price Lock] = @PriceLock,
                        [Price 1] = @Price1,
                        [Price 2] = @Price2,
                        [Price 3] = @Price3,
                        [Price 4] = @Price4,
                        [Price 5] = @Price5,

                        [Shift Event 1] = @ShiftEvent1,
                        [Shift Event 2] = @ShiftEvent2,
                        [Shift Event 3] = @ShiftEvent3,
                        [Shift Event 4] = @ShiftEvent4,
                        [Shift Event 5] = @ShiftEvent5,
                        [Shift Price 1] = @ShiftPrice1,
                        [Shift Price 2] = @ShiftPrice2,
                        [Shift Price 3] = @ShiftPrice3,
                        [Shift Price 4] = @ShiftPrice4,
                        [Shift Price 5] = @ShiftPrice5,
                        [Shift Mod 1] = @ShiftMod1,
                        [Shift Mod 2] = @ShiftMod2,
                        [Shift Mod 3] = @ShiftMod3,
                        [Shift Mod 4] = @ShiftMod4,
                        [Shift Mod 5] = @ShiftMod5,
                        [EModShift 1] = @EModShift1,
                        [EModShift 2] = @EModShift2,
                        [EModShift 3] = @EModShift3,
                        [EModShift 4] = @EModShift4,
                        [EModShift 5] = @EModShift5,
                        [PModShift 1] = @PModShift1,
                        [PModShift 2] = @PModShift2,
                        [PModShift 3] = @PModShift3,
                        [PModShift 4] = @PModShift4,
                        [PModShift 5] = @PModShift5,
                        [Shift 1 PLU] = @ShiftPLU1,
                        [Shift 2 PLU] = @ShiftPLU2,
                        [Shift 3 PLU] = @ShiftPLU3,
                        [Shift 4 PLU] = @ShiftPLU4,
                        [Shift 5 PLU] = @ShiftPLU5,

                        [PLU] = -1
                    WHERE [Menu Item Name] = @OriginalMenuItemName", cn, tran);

                        AddParameters(update, model);
                        update.Parameters.AddWithValue("@OriginalMenuItemName", model.OriginalMenuItemName.Trim());

                        rows = update.ExecuteNonQuery();
                    }

                    if (rows != 1)
                    {
                        tran.Rollback();
                        return Json(new { success = false, message = "Save failed. Item was not found." });
                    }

                    string savedName = model.MenuItemName.Trim();

                    using var deleteMods = new SqlCommand(@"
                DELETE FROM [Menu Mods]
                WHERE [Menu Item Name] = @MenuItemName", cn, tran);

                    deleteMods.Parameters.AddWithValue("@MenuItemName", savedName);
                    deleteMods.ExecuteNonQuery();

                    if (model.MinMaxChoices)
                    {
                        InsertMenuMod(cn, tran, savedName, model.Modifier1, model.ModifierMin1, model.ModifierMax1);
                        InsertMenuMod(cn, tran, savedName, model.Modifier2, model.ModifierMin2, model.ModifierMax2);
                        InsertMenuMod(cn, tran, savedName, model.Modifier3, model.ModifierMin3, model.ModifierMax3);
                        InsertMenuMod(cn, tran, savedName, model.Modifier4, model.ModifierMin4, model.ModifierMax4);
                        InsertMenuMod(cn, tran, savedName, model.Modifier5, model.ModifierMin5, model.ModifierMax5);
                        InsertMenuMod(cn, tran, savedName, model.Modifier6, model.ModifierMin6, model.ModifierMax6);
                    }

                    tran.Commit();

                    _terminals.FlagAllTerminals();  // tell POS terminals to reload

                    return Json(new
                    {
                        success = true,
                        message = model.IsNewItem ? "New menu item added." : "Menu item saved."
                    });
                }
                catch
                {
                    tran.Rollback();
                    throw;
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public IActionResult GetDepartments()
        {
            var list = new List<string>();

            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();

            using var cmd = new SqlCommand(@"
                SELECT [DepartmentName]
                FROM [Departments]
                WHERE [DepartmentName] IS NOT NULL
                  AND LTRIM(RTRIM([DepartmentName])) <> ''
                ORDER BY [DepartmentName]", cn);

            using var dr = cmd.ExecuteReader();

            while (dr.Read())
                list.Add(dr["DepartmentName"].ToString() ?? "");

            return Json(list);
        }

        [HttpGet]
        public IActionResult GetReportCategories()
        {
            var list = new List<string>();

            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();

            using var cmd = new SqlCommand(@"
                SELECT [Report Name]
                FROM [Report Categories]
                WHERE [Report Name] IS NOT NULL
                  AND LTRIM(RTRIM([Report Name])) <> ''
                ORDER BY [Report Name]", cn);

            using var dr = cmd.ExecuteReader();

            while (dr.Read())
                list.Add(dr["Report Name"].ToString() ?? "");

            return Json(list);
        }

        [HttpGet]
        public IActionResult GetSystemPrinters()
        {
            var list = new List<string>();

            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();

            using var cmd = new SqlCommand(@"
        SELECT DISTINCT 
            LTRIM(RTRIM(ISNULL(CAST([Printer Name] AS varchar(100)), ''))) AS PrinterName
        FROM [System Printers]
        WHERE LTRIM(RTRIM(ISNULL(CAST([Printer Name] AS varchar(100)), ''))) <> ''
        ORDER BY PrinterName", cn);

            using var dr = cmd.ExecuteReader();

            while (dr.Read())
                list.Add(dr["PrinterName"].ToString() ?? "");

            return Json(list);
        }
        private void AddParameters(SqlCommand cmd, MenuItemBasicDto model)
        {
            cmd.Parameters.AddWithValue("@MenuItemName", model.MenuItemName.Trim());
            cmd.Parameters.AddWithValue("@DepartmentName", model.DepartmentName ?? "");
            cmd.Parameters.AddWithValue("@ReportCategory", model.ReportCategory ?? "");

            cmd.Parameters.AddWithValue("@RegularPrice", model.RegularPrice);
            cmd.Parameters.AddWithValue("@InventoryAmount", model.InventoryAmount);
            cmd.Parameters.AddWithValue("@ProductCost", model.ProductCost);
            cmd.Parameters.AddWithValue("@UpCharge", model.UpCharge);

            cmd.Parameters.AddWithValue("@BarCode", model.BarCode ?? "");
            cmd.Parameters.AddWithValue("@ItemInfo", model.ItemInfo ?? "");

            cmd.Parameters.AddWithValue("@SalesTaxTable", model.SalesTaxTable);
            cmd.Parameters.AddWithValue("@SpecialTaxTable", model.SpecialTaxTable);
            cmd.Parameters.AddWithValue("@LiquorTaxAmount", model.LiquorTaxAmount);

            cmd.Parameters.AddWithValue("@Remote1", model.Remote1 ?? "");
            cmd.Parameters.AddWithValue("@Remote2", model.Remote2 ?? "");
            cmd.Parameters.AddWithValue("@Remote3", model.Remote3 ?? "");
            cmd.Parameters.AddWithValue("@Remote4", model.Remote4 ?? "");
            cmd.Parameters.AddWithValue("@Remote5", model.Remote5 ?? "");
            cmd.Parameters.AddWithValue("@Remote6", model.Remote6 ?? "");

            cmd.Parameters.AddWithValue("@Avalibility", model.Avalibility);
            cmd.Parameters.AddWithValue("@Taxable", model.Taxable);
            cmd.Parameters.AddWithValue("@UseModifiers", model.UseModifiers);
            cmd.Parameters.AddWithValue("@ModifiersAreOptional", model.ModifiersAreOptional);
            cmd.Parameters.AddWithValue("@MinMaxChoices", model.MinMaxChoices);
            cmd.Parameters.AddWithValue("@OpenPrice", model.OpenPrice);
            cmd.Parameters.AddWithValue("@QuickPick", model.QuickPick);
            cmd.Parameters.AddWithValue("@QuickButton", model.QuickButton);
            cmd.Parameters.AddWithValue("@OnLine", model.OnLine);

            cmd.Parameters.AddWithValue("@TimedPrice", model.TimedPrice);

            cmd.Parameters.AddWithValue("@TimeEvent1", model.TimeEvent1 ?? "");
            cmd.Parameters.AddWithValue("@TimeEvent2", model.TimeEvent2 ?? "");
            cmd.Parameters.AddWithValue("@TimeEvent3", model.TimeEvent3 ?? "");
            cmd.Parameters.AddWithValue("@TimeEvent4", model.TimeEvent4 ?? "");
            cmd.Parameters.AddWithValue("@TimeEvent5", model.TimeEvent5 ?? "");

            cmd.Parameters.AddWithValue("@TimePrice1", model.TimePrice1);
            cmd.Parameters.AddWithValue("@TimePrice2", model.TimePrice2);
            cmd.Parameters.AddWithValue("@TimePrice3", model.TimePrice3);
            cmd.Parameters.AddWithValue("@TimePrice4", model.TimePrice4);
            cmd.Parameters.AddWithValue("@TimePrice5", model.TimePrice5);

            cmd.Parameters.AddWithValue("@Modifier1", model.Modifier1 ?? "");
            cmd.Parameters.AddWithValue("@Modifier2", model.Modifier2 ?? "");
            cmd.Parameters.AddWithValue("@Modifier3", model.Modifier3 ?? "");
            cmd.Parameters.AddWithValue("@Modifier4", model.Modifier4 ?? "");
            cmd.Parameters.AddWithValue("@Modifier5", model.Modifier5 ?? "");
            cmd.Parameters.AddWithValue("@Modifier6", model.Modifier6 ?? "");

            cmd.Parameters.AddWithValue("@ModifierOption1", model.ModifierOption1);
            cmd.Parameters.AddWithValue("@ModifierOption2", model.ModifierOption2);
            cmd.Parameters.AddWithValue("@ModifierOption3", model.ModifierOption3);
            cmd.Parameters.AddWithValue("@ModifierOption4", model.ModifierOption4);
            cmd.Parameters.AddWithValue("@ModifierOption5", model.ModifierOption5);
            cmd.Parameters.AddWithValue("@ModifierOption6", model.ModifierOption6);

            cmd.Parameters.AddWithValue("@PriceLevels", model.PriceLevels);
            cmd.Parameters.AddWithValue("@PriceLock", model.PriceLock);
            cmd.Parameters.AddWithValue("@Price1", model.Price1);
            cmd.Parameters.AddWithValue("@Price2", model.Price2);
            cmd.Parameters.AddWithValue("@Price3", model.Price3);
            cmd.Parameters.AddWithValue("@Price4", model.Price4);
            cmd.Parameters.AddWithValue("@Price5", model.Price5);

            cmd.Parameters.AddWithValue("@ShiftEvent1", model.ShiftEvent1 ?? "");
            cmd.Parameters.AddWithValue("@ShiftEvent2", model.ShiftEvent2 ?? "");
            cmd.Parameters.AddWithValue("@ShiftEvent3", model.ShiftEvent3 ?? "");
            cmd.Parameters.AddWithValue("@ShiftEvent4", model.ShiftEvent4 ?? "");
            cmd.Parameters.AddWithValue("@ShiftEvent5", model.ShiftEvent5 ?? "");
            cmd.Parameters.AddWithValue("@ShiftPrice1", model.ShiftPrice1);
            cmd.Parameters.AddWithValue("@ShiftPrice2", model.ShiftPrice2);
            cmd.Parameters.AddWithValue("@ShiftPrice3", model.ShiftPrice3);
            cmd.Parameters.AddWithValue("@ShiftPrice4", model.ShiftPrice4);
            cmd.Parameters.AddWithValue("@ShiftPrice5", model.ShiftPrice5);
            cmd.Parameters.AddWithValue("@ShiftMod1", model.ShiftMod1 ?? "");
            cmd.Parameters.AddWithValue("@ShiftMod2", model.ShiftMod2 ?? "");
            cmd.Parameters.AddWithValue("@ShiftMod3", model.ShiftMod3 ?? "");
            cmd.Parameters.AddWithValue("@ShiftMod4", model.ShiftMod4 ?? "");
            cmd.Parameters.AddWithValue("@ShiftMod5", model.ShiftMod5 ?? "");
            cmd.Parameters.AddWithValue("@EModShift1", model.EModShift1);
            cmd.Parameters.AddWithValue("@EModShift2", model.EModShift2);
            cmd.Parameters.AddWithValue("@EModShift3", model.EModShift3);
            cmd.Parameters.AddWithValue("@EModShift4", model.EModShift4);
            cmd.Parameters.AddWithValue("@EModShift5", model.EModShift5);
            cmd.Parameters.AddWithValue("@PModShift1", model.PModShift1);
            cmd.Parameters.AddWithValue("@PModShift2", model.PModShift2);
            cmd.Parameters.AddWithValue("@PModShift3", model.PModShift3);
            cmd.Parameters.AddWithValue("@PModShift4", model.PModShift4);
            cmd.Parameters.AddWithValue("@PModShift5", model.PModShift5);
            cmd.Parameters.AddWithValue("@ShiftPLU1", model.ShiftPLU1);
            cmd.Parameters.AddWithValue("@ShiftPLU2", model.ShiftPLU2);
            cmd.Parameters.AddWithValue("@ShiftPLU3", model.ShiftPLU3);
            cmd.Parameters.AddWithValue("@ShiftPLU4", model.ShiftPLU4);
            cmd.Parameters.AddWithValue("@ShiftPLU5", model.ShiftPLU5);
        }

        private void InsertMenuMod(
            SqlConnection cn,
            SqlTransaction tran,
            string menuItemName,
            string modifierName,
            int min,
            int max)
        {
            if (string.IsNullOrWhiteSpace(modifierName))
                return;

            using var cmd = new SqlCommand(@"
        INSERT INTO [Menu Mods]
        (
            [Menu Item Name],
            [Modifier Name],
            [Min],
            [Max]
        )
        VALUES
        (
            @MenuItemName,
            @ModifierName,
            @Min,
            @Max
        )", cn, tran);

            cmd.Parameters.AddWithValue("@MenuItemName", menuItemName);
            cmd.Parameters.AddWithValue("@ModifierName", modifierName.Trim());
            cmd.Parameters.AddWithValue("@Min", min);
            cmd.Parameters.AddWithValue("@Max", max);

            cmd.ExecuteNonQuery();
        }

        [HttpGet]
        public IActionResult GetTimedEvents()
        {
            var list = new List<string>();

            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();

            using var cmd = new SqlCommand(@"
        SELECT [Event Name]
        FROM [Timed Pricing]
        WHERE [Event Name] IS NOT NULL
          AND LTRIM(RTRIM([Event Name])) <> ''
        ORDER BY [Event Name]", cn);

            using var dr = cmd.ExecuteReader();

            while (dr.Read())
                list.Add(dr["Event Name"].ToString() ?? "");

            return Json(list);
        }

        [HttpGet]
        public IActionResult GetModifierNames()
        {
            var list = new List<string>();

            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();

            using var cmd = new SqlCommand(@"
        SELECT [Modifiers]
        FROM [Modifier Names]
        WHERE [Modifiers] IS NOT NULL
          AND LTRIM(RTRIM([Modifiers])) <> ''
        ORDER BY [Modifiers]", cn);

            using var dr = cmd.ExecuteReader();

            while (dr.Read())
                list.Add(dr["Modifiers"].ToString() ?? "");

            return Json(list);
        }
    }
}
