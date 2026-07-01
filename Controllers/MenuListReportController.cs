using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    // Web port of the legacy VB "PrintStoreMenu_Click" report (landscape view mode).
    // Lists menu items from [MenuItemCountDown] where [Quick Pick] = 0,
    // grouped/ordered by Report Category then item name.
    public class MenuListReportController : Controller
    {
        private readonly IConfiguration _cfg;
        public MenuListReportController(IConfiguration cfg) => _cfg = cfg;

        private SqlConnection Cn() => new(_cfg.GetConnectionString("DefaultConnection"));

        [HttpGet]
        public IActionResult Index() =>
            View("~/Views/MSaleAdmin/MenuListReport.cshtml");

        [HttpGet]
        public IActionResult GetMenuList()
        {
            try
            {
                using var cn = Cn();
                cn.Open();

                const string sql = @"
                    SELECT *
                    FROM [MSaleData].[dbo].[MenuItemCountDown]
                    WHERE ISNULL([Quick Pick],0) = 0
                    ORDER BY [Report Category], [Menu Item Name]";

                using var cmd = new SqlCommand(sql, cn);
                using var rd = cmd.ExecuteReader();

                var rows = new List<object>();
                while (rd.Read())
                {
                    rows.Add(new
                    {
                        category = S(rd, "Report Category"),
                        itemName = S(rd, "Menu Item Name").Trim(),
                        plu = S(rd, "Bar Code"),
                        price = Dec(rd, "Regular price"),
                        taxable = Flag(rd, "Taxable"),
                        mod1 = S(rd, "Modifier 1").Trim(),
                        mod2 = S(rd, "Modifier 2").Trim(),
                        mod3 = S(rd, "Modifier 3").Trim(),
                        mod4 = S(rd, "Modifier 4").Trim(),
                        mod5 = S(rd, "Modifier 5").Trim(),
                        mod6 = S(rd, "Modifier 6").Trim(),
                        rmt1 = S(rd, "Remote 1").Trim(),
                        rmt2 = S(rd, "Remote 2").Trim()
                    });
                }

                return Json(new { success = true, data = rows });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        private static bool HasColumn(SqlDataReader rd, string name)
        {
            for (int i = 0; i < rd.FieldCount; i++)
                if (string.Equals(rd.GetName(i), name, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        private static string S(SqlDataReader rd, string col) =>
            HasColumn(rd, col) && rd[col] != DBNull.Value ? rd[col].ToString() ?? "" : "";

        private static decimal Dec(SqlDataReader rd, string col) =>
            HasColumn(rd, col) && rd[col] != DBNull.Value && decimal.TryParse(rd[col].ToString(), out var d) ? d : 0;

        private static bool Flag(SqlDataReader rd, string col)
        {
            if (!HasColumn(rd, col) || rd[col] == DBNull.Value) return false;
            var s = rd[col].ToString();
            return s == "1" || string.Equals(s, "true", StringComparison.OrdinalIgnoreCase);
        }
    }
}
