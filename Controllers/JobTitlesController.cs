using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/job-titles")]
    public class JobTitlesController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public JobTitlesController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetJobTitles()
        {
            try
            {
                string cs =
                    _configuration.GetConnectionString(
                        "DefaultConnection");

                var list =
                    new List<JobTitleDto>();

                using var cn =
                    new SqlConnection(cs);

                cn.Open();

                string sql = @"
SELECT
 [Job Title],
 [Tipped Position],
 [Tip Pool Percent],
 [Job Code Number],
 [Drawer]
FROM [dbo].[Job Titles]
ORDER BY [Job Title]";

                using var cmd =
                    new SqlCommand(sql, cn);

                using var rdr =
                    cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(
                     new JobTitleDto
                     {
                         JobTitle =
                        GetStr(rdr, "Job Title"),

                         TippedPosition =
                        ToBool(
                          rdr["Tipped Position"]),

                         TipPoolPercent =
                        GetDecimal(
                          rdr, "Tip Pool Percent"),

                         JobCodeNumber =
                        GetStr(
                          rdr, "Job Code Number"),

                         Drawer =
                        GetStr(
                          rdr, "Drawer")
                     });
                }

                return Ok(list);
            }
            catch (Exception ex)
            {
                return StatusCode(
                 500,
                 "Error: " + ex.Message);
            }
        }

        private static string GetStr(
         SqlDataReader rdr,
         string name)
        {
            return rdr[name] == DBNull.Value
              ? ""
              : rdr[name]?.ToString() ?? "";
        }

        private static decimal GetDecimal(
          SqlDataReader rdr,
          string name)
        {
            if (rdr[name] == DBNull.Value)
                return 0m;

            decimal.TryParse(
             rdr[name]?.ToString(),
             out decimal value);

            return value;
        }

        private static bool ToBool(
          object value)
        {
            if (value == null ||
               value == DBNull.Value)
                return false;

            string s =
             value.ToString()?
             .Trim()
             .ToLower() ?? "";

            return s == "1" ||
                   s == "true" ||
                   s == "yes" ||
                   s == "y";
        }
    }

    public class JobTitleDto
    {
        public string JobTitle { get; set; } = "";
        public bool TippedPosition { get; set; }
        public decimal TipPoolPercent { get; set; }
        public string JobCodeNumber { get; set; } = "";
        public string Drawer { get; set; } = "";
    }
}