using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/menu-items")]
    public class MenuItemsController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public MenuItemsController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetMenuItems()
        {
            List<string> list = new List<string>();

            using (SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(@"
            SELECT [Menu Item]
            FROM [Menu Items]
            WHERE [Menu Item] IS NOT NULL
            ORDER BY [Menu Item]", cn);

                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        list.Add(dr["Menu Item"].ToString());
                    }
                }
            }

            return new JsonResult(list);
        }

        private static string GetStr(SqlDataReader rdr, string name)
        {
            return rdr[name] == DBNull.Value ? "" : rdr[name]?.ToString() ?? "";
        }

        private static int GetInt(SqlDataReader rdr, string name)
        {
            if (rdr[name] == DBNull.Value) return 0;
            int.TryParse(rdr[name]?.ToString(), out int value);
            return value;
        }

        private static decimal GetDecimal(SqlDataReader rdr, string name)
        {
            if (rdr[name] == DBNull.Value) return 0m;
            decimal.TryParse(rdr[name]?.ToString(), out decimal value);
            return value;
        }
    }

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
}