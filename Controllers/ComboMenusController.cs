using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/combo-menus")]
    public class ComboMenusController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public ComboMenusController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetComboMenus()
        {
            try
            {
                string cs = _configuration.GetConnectionString("DefaultConnection");

                var list = new List<ComboMenuDto>();

                using var cn = new SqlConnection(cs);
                cn.Open();

                string sql = @"
SELECT
   [Combo Name],
   [POS],
   [Menu Item]
FROM [dbo].[Combo Menu]
ORDER BY [Combo Name], [POS]";

                using var cmd = new SqlCommand(sql, cn);
                using var rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(new ComboMenuDto
                    {
                        ComboName = rdr["Combo Name"]?.ToString() ?? "",
                        Position = rdr["POS"] != DBNull.Value
                                  ? Convert.ToInt32(rdr["POS"])
                                  : 0,
                        MenuItemName = rdr["Menu Item"]?.ToString() ?? ""
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

    public class ComboMenuDto
    {
        public string ComboName { get; set; } = "";
        public int Position { get; set; }
        public string MenuItemName { get; set; } = "";
    }
}