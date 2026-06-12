using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/menu-button-layout")]
    public class MenuButtonLayoutController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public MenuButtonLayoutController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetLayout()
        {
            try
            {
                string cs = _configuration.GetConnectionString("DefaultConnection");

                var list = new List<MenuButtonLayoutDto>();

                using var cn = new SqlConnection(cs);
                cn.Open();

                string sql = @"
SELECT
    [Menu Item],
    [Category],
    [ForeCol],
    [BackCol],
    [Icon],
    [Width],
    [Height],
    [FontSize],
    [Left],
    [Top],
    [Button Assigned],
    [Sound]
FROM [dbo].[Button Layout]";

                using var cmd = new SqlCommand(sql, cn);
                using var rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(new MenuButtonLayoutDto
                    {
                        MenuItemName = rdr["Menu Item"]?.ToString() ?? "",
                        CategoryName = rdr["Category"]?.ToString() ?? "",
                        ForeColor = rdr["ForeCol"]?.ToString() ?? "",
                        BackColor = rdr["BackCol"]?.ToString() ?? "",
                        Icon = rdr["Icon"]?.ToString() ?? "",
                        Width = rdr["Width"] != DBNull.Value ? Convert.ToInt32(rdr["Width"]) : 0,
                        Height = rdr["Height"] != DBNull.Value ? Convert.ToInt32(rdr["Height"]) : 0,
                        FontSize = rdr["FontSize"] != DBNull.Value ? Convert.ToInt32(rdr["FontSize"]) : 0,
                        Left = rdr["Left"] != DBNull.Value ? Convert.ToInt32(rdr["Left"]) : 0,
                        Top = rdr["Top"] != DBNull.Value ? Convert.ToInt32(rdr["Top"]) : 0,
                        ButtonAssigned = rdr["Button Assigned"]?.ToString() ?? "",
                        Sound = rdr["Sound"]?.ToString() ?? ""
                    });
                }

                return Ok(list);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Error: " + ex.Message);
            }
        }
    }

    public class MenuButtonLayoutDto
    {
        public string MenuItemName { get; set; } = "";
        public string CategoryName { get; set; } = "";
        public string ForeColor { get; set; } = "";
        public string BackColor { get; set; } = "";
        public string Icon { get; set; } = "";
        public int Width { get; set; }
        public int Height { get; set; }
        public int FontSize { get; set; }
        public int Left { get; set; }
        public int Top { get; set; }
        public string ButtonAssigned { get; set; } = "";
        public string Sound { get; set; } = "";
    }
}