using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/departments")]
    public class DepartmentsController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public DepartmentsController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetDepartments()
        {
            try
            {
                string cs = _configuration.GetConnectionString("DefaultConnection");

                if (string.IsNullOrWhiteSpace(cs))
                {
                    return StatusCode(500, "Error: DefaultConnection is missing or empty in appsettings.json");
                }

                var list = new List<DepartmentDto>();

                using var cn = new SqlConnection(cs);
                cn.Open();

                string sql = @"
SELECT
    [Report Name],
    [Flag],
    [Record Number],
    [Remote 1],
    [Remote 2],
    [Remote 3],
    [Remote 4],
    [Remote 5],
    [Remote 6],
    [Master Name],
    [Timed Event 1],
    [Timed Event 2],
    [Timed Event 3],
    [Timed Event 4],
    [Timed Event 5],
    [Department No]
FROM [dbo].[Report Categories]
ORDER BY [Report Name]";

                using var cmd = new SqlCommand(sql, cn);
                using var rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(new DepartmentDto
                    {
                        DepartmentId = rdr["Record Number"] != DBNull.Value
                            ? Convert.ToInt32(rdr["Record Number"])
                            : 0,

                        DepartmentName = rdr["Report Name"]?.ToString() ?? "",

                        DepartmentCode = rdr["Department No"]?.ToString() ?? "",

                        ParentName = rdr["Master Name"]?.ToString() ?? "",

                        IsActive = rdr["Flag"] != DBNull.Value &&
                                   Convert.ToInt32(rdr["Flag"]) != 0,

                        Remote1 = rdr["Remote 1"]?.ToString() ?? "",
                        Remote2 = rdr["Remote 2"]?.ToString() ?? "",
                        Remote3 = rdr["Remote 3"]?.ToString() ?? "",
                        Remote4 = rdr["Remote 4"]?.ToString() ?? "",
                        Remote5 = rdr["Remote 5"]?.ToString() ?? "",
                        Remote6 = rdr["Remote 6"]?.ToString() ?? "",

                        TimedEvent1 = rdr["Timed Event 1"]?.ToString() ?? "",
                        TimedEvent2 = rdr["Timed Event 2"]?.ToString() ?? "",
                        TimedEvent3 = rdr["Timed Event 3"]?.ToString() ?? "",
                        TimedEvent4 = rdr["Timed Event 4"]?.ToString() ?? "",
                        TimedEvent5 = rdr["Timed Event 5"]?.ToString() ?? ""
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

    public class DepartmentDto
    {
        public int DepartmentId { get; set; }
        public string DepartmentName { get; set; } = "";
        public string DepartmentCode { get; set; } = "";
        public string ParentName { get; set; } = "";
        public bool IsActive { get; set; }

        public string Remote1 { get; set; } = "";
        public string Remote2 { get; set; } = "";
        public string Remote3 { get; set; } = "";
        public string Remote4 { get; set; } = "";
        public string Remote5 { get; set; } = "";
        public string Remote6 { get; set; } = "";

        public string TimedEvent1 { get; set; } = "";
        public string TimedEvent2 { get; set; } = "";
        public string TimedEvent3 { get; set; } = "";
        public string TimedEvent4 { get; set; } = "";
        public string TimedEvent5 { get; set; } = "";
    }
}