using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/daily-specials")]
    public class DailySpecialsController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public DailySpecialsController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetDailySpecials()
        {
            try
            {
                string cs =
                    _configuration.GetConnectionString(
                        "DefaultConnection");

                var list =
                    new List<DailySpecialDto>();

                using var cn =
                    new SqlConnection(cs);

                cn.Open();

                string sql = @"
SELECT
 [Special 1],
 [Special 2],
 [Special 3],
 [Special 4],
 [Flash Mess 1],
 [Flash Mess 2],
 [Flash Mess 3],
 [Flash Mess 4],
 [Day]
FROM [dbo].[Daily Specials]
ORDER BY [Day]";

                using var cmd =
                    new SqlCommand(sql, cn);

                using var rdr =
                    cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(
                        new DailySpecialDto
                        {
                            Day =
                              GetStr(rdr, "Day"),

                            Special1 =
                              GetStr(rdr, "Special 1"),

                            Special2 =
                              GetStr(rdr, "Special 2"),

                            Special3 =
                              GetStr(rdr, "Special 3"),

                            Special4 =
                              GetStr(rdr, "Special 4"),

                            FlashMess1 =
                              GetStr(rdr, "Flash Mess 1"),

                            FlashMess2 =
                              GetStr(rdr, "Flash Mess 2"),

                            FlashMess3 =
                              GetStr(rdr, "Flash Mess 3"),

                            FlashMess4 =
                              GetStr(rdr, "Flash Mess 4")
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
    }

    public class DailySpecialDto
    {
        public string Day { get; set; } = ""
;
        public string Special1 { get; set; } = "";
        public string Special2 { get; set; } = "";
        public string Special3 { get; set; } = "";
        public string Special4 { get; set; } = "";

        public string FlashMess1 { get; set; } = "";
        public string FlashMess2 { get; set; } = "";
        public string FlashMess3 { get; set; } = "";
        public string FlashMess4 { get; set; } = "";
    }
}