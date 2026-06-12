using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MSaleWebServer.Models;

namespace MSaleWebServer.Controllers
{
    public class MenuItemController : Controller
    {
        private readonly IConfiguration _configuration;

        public MenuItemController(IConfiguration configuration)
        {
            _configuration = configuration;
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
                    ISNULL([On Line], 0) AS [On Line]
                FROM [Menu Items]
                WHERE [Menu Item Name] = @MenuItemName", cn);

            cmd.Parameters.AddWithValue("@MenuItemName", menuItemName ?? "");

            using var dr = cmd.ExecuteReader();

            if (!dr.Read())
                return Json(new { success = false, message = "Menu item not found." });

            var item = new MenuItemBasicDto
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
                LiquorTaxAmount = Convert.ToInt32(dr["Liquor Tax Amount"]),

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
                OnLine = Convert.ToBoolean(dr["On Line"])
            };

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

                if (model.IsNewItem)
                {
                    using var check = new SqlCommand(@"
                        SELECT COUNT(*)
                        FROM [Menu Items]
                        WHERE [Menu Item Name] = @MenuItemName", cn);

                    check.Parameters.AddWithValue("@MenuItemName", model.MenuItemName.Trim());

                    int exists = Convert.ToInt32(check.ExecuteScalar());

                    if (exists > 0)
                        return Json(new { success = false, message = "That menu item already exists." });

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
                            [Open Price],
                            [Quick Pick],
                            [Quick Button],
                            [On Line],
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
                            @OpenPrice,
                            @QuickPick,
                            @QuickButton,
                            @OnLine,
                            ISNULL((SELECT MAX([Record Number]) + 1 FROM [Menu Items]), 1)
                        )", cn);

                    AddParameters(insert, model);

                    int rows = insert.ExecuteNonQuery();

                    return Json(new
                    {
                        success = rows == 1,
                        message = rows == 1 ? "New menu item added." : "Insert failed."
                    });
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(model.OriginalMenuItemName))
                        return Json(new { success = false, message = "Original menu item name is missing." });

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
                            [Open Price] = @OpenPrice,
                            [Quick Pick] = @QuickPick,
                            [Quick Button] = @QuickButton,
                            [On Line] = @OnLine
                        WHERE [Menu Item Name] = @OriginalMenuItemName", cn);

                    AddParameters(update, model);
                    update.Parameters.AddWithValue("@OriginalMenuItemName", model.OriginalMenuItemName.Trim());

                    int rows = update.ExecuteNonQuery();

                    return Json(new
                    {
                        success = rows == 1,
                        message = rows == 1 ? "Menu item saved." : "Update failed. Item was not found."
                    });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        private static void AddParameters(SqlCommand cmd, MenuItemBasicDto model)
        {
            cmd.Parameters.AddWithValue("@MenuItemName", model.MenuItemName?.Trim() ?? "");
            cmd.Parameters.AddWithValue("@DepartmentName", model.DepartmentName?.Trim() ?? "");
            cmd.Parameters.AddWithValue("@ReportCategory", model.ReportCategory?.Trim() ?? "");

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
            cmd.Parameters.AddWithValue("@OpenPrice", model.OpenPrice);
            cmd.Parameters.AddWithValue("@QuickPick", model.QuickPick);
            cmd.Parameters.AddWithValue("@QuickButton", model.QuickButton);
            cmd.Parameters.AddWithValue("@OnLine", model.OnLine);
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
                SELECT DISTINCT [Printer Name]
                FROM [System Printers]
                WHERE [Printer Name] IS NOT NULL
                  AND LTRIM(RTRIM([Printer Name])) <> ''
                ORDER BY [Printer Name]", cn);

            using var dr = cmd.ExecuteReader();

            while (dr.Read())
                list.Add(dr["Printer Name"].ToString() ?? "");

            return Json(list);
        }
    }
}