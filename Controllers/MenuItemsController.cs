using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    /// <summary>
    /// Full menu-item feed for the StoreSyncAgent (down-sync to POS stores).
    /// </summary>
    [ApiController]
    [Route("api/menu-items")]
    public class MenuItemsController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        public MenuItemsController(IConfiguration configuration) => _configuration = configuration;

        public class MenuItemDto
        {
            public string MenuItemName { get; set; } = "";
            public string DepartmentName { get; set; } = "";
            public decimal RegularPrice { get; set; }
            public decimal InventoryAmount { get; set; }
            public string Availability { get; set; } = "";
            public string Taxable { get; set; } = "";
            public string RemotePrinting { get; set; } = "";
            public string SalesTaxTable { get; set; } = "";
            public string SpecialTaxTable { get; set; } = "";
            public decimal LiquorTaxAmount { get; set; }
            public string UseModifiers { get; set; } = "";
            public string PLU { get; set; } = "";
            public int RecordNumber { get; set; }
            public string ReportCategory { get; set; } = "";
            public string BarCode { get; set; } = "";
            public string OpenPrice { get; set; } = "";
            public decimal ProductCost { get; set; }
            public string ManagerAccess { get; set; } = "";
            public string QuickPick { get; set; } = "";
            public string Online { get; set; } = "";
        }

        [HttpGet]
        public IActionResult GetMenuItems()
        {
            var list = new List<MenuItemDto>();
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();

            // NB: [Avalibility] spelling matches the legacy table
            const string sql = @"
SELECT [Menu Item Name],[Department Name],[Regular Price],[Inventory Amount],[Avalibility],
       [Taxable],[Remote Printing],[Sales Tax Table],[Special Tax Table],[Liquor Tax Amount],
       [Use Modifiers],[PLU],[Record Number],[Report Category],[Bar Code],[Open Price],
       [Product Cost],[Manager Access],[Quick Pick]
FROM [Menu Items]
WHERE [Menu Item Name] IS NOT NULL AND LTRIM(RTRIM([Menu Item Name])) <> ''
ORDER BY [Menu Item Name]";

            static string S(SqlDataReader r, int i) => r.IsDBNull(i) ? "" : Convert.ToString(r.GetValue(i))!.Trim();
            static decimal D(SqlDataReader r, int i) { try { return r.IsDBNull(i) ? 0m : Convert.ToDecimal(r.GetValue(i)); } catch { return 0m; } }
            static int I(SqlDataReader r, int i) { try { return r.IsDBNull(i) ? 0 : Convert.ToInt32(r.GetValue(i)); } catch { return 0; } }

            using var cmd = new SqlCommand(sql, cn);
            using var dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                list.Add(new MenuItemDto
                {
                    MenuItemName    = S(dr, 0),  DepartmentName = S(dr, 1),
                    RegularPrice    = D(dr, 2),  InventoryAmount = D(dr, 3),
                    Availability    = S(dr, 4),  Taxable = S(dr, 5),
                    RemotePrinting  = S(dr, 6),  SalesTaxTable = S(dr, 7),
                    SpecialTaxTable = S(dr, 8),  LiquorTaxAmount = D(dr, 9),
                    UseModifiers    = S(dr, 10), PLU = S(dr, 11),
                    RecordNumber    = I(dr, 12), ReportCategory = S(dr, 13),
                    BarCode         = S(dr, 14), OpenPrice = S(dr, 15),
                    ProductCost     = D(dr, 16), ManagerAccess = S(dr, 17),
                    QuickPick       = S(dr, 18), Online = ""
                });
            }
            return Ok(list);
        }
    }
}
