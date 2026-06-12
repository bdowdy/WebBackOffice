using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;


namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/menuaccess")]
    public class MenuAccessController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public MenuAccessController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetMenuAccess()
        {
            try
            {
                string cs = _configuration.GetConnectionString("DefaultConnection");

                if (string.IsNullOrWhiteSpace(cs))
                {
                    return StatusCode(500, "Error: DefaultConnection is missing or empty in appsettings.json");
                }

                var list = new List<MenuAccessDto>();

                using var cn = new SqlConnection(cs);
                cn.Open();

                string sql = @"
SELECT
    [Menu Name],
    [Access Required]
FROM [dbo].[Menu Access]
ORDER BY [Menu Name]";

                using var cmd = new SqlCommand(sql, cn);
                using var rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(new MenuAccessDto
                    {
                        MenuName = rdr["Menu Name"]?.ToString() ?? "",
                        AccessRequired = rdr["Access Required"] != DBNull.Value
                            ? Convert.ToInt32(rdr["Access Required"])
                            : 0
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

    public class MenuAccessDto
    {
        public string MenuName { get; set; } = "";
        public int AccessRequired { get; set; }
    }
}