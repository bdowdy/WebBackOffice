using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/accesslevels")]
    public class AccessLevelsController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public AccessLevelsController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetAccessLevels()
        {
            try
            {
                string cs = _configuration.GetConnectionString("DefaultConnection");

                if (string.IsNullOrWhiteSpace(cs))
                {
                    return StatusCode(500, "Connection string missing");
                }

                var list = new List<AccessLevelDto>();

                using var cn = new SqlConnection(cs);
                cn.Open();

                string sql = @"
SELECT 
    [Access Name],
    [Access Level]
FROM [dbo].[Access Levels]
ORDER BY [Access Level]";

                using var cmd = new SqlCommand(sql, cn);
                using var rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(new AccessLevelDto
                    {
                        AccessName = rdr["Access Name"]?.ToString() ?? "",
                        AccessLevel = rdr["Access Level"] != DBNull.Value
                            ? Convert.ToInt32(rdr["Access Level"])
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

    public class AccessLevelDto
    {
        public string AccessName { get; set; } = "";
        public int AccessLevel { get; set; }
    }
}