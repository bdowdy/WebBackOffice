using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/discount-events")]
    public class DiscountEventsController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public DiscountEventsController(
            IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetDiscountEvents()
        {
            try
            {
                string cs =
                    _configuration.GetConnectionString(
                        "DefaultConnection");

                var list =
                    new List<DiscountEventDto>();

                using var cn =
                    new SqlConnection(cs);

                cn.Open();

                string sql = @"
SELECT
   [Event Name],
   [POS],
   [Discount Name]
FROM [dbo].[Discount Events]
ORDER BY [Event Name],[POS]";

                using var cmd =
                    new SqlCommand(sql, cn);

                using var rdr =
                    cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(
                     new DiscountEventDto
                     {
                         EventName =
                        GetStr(rdr, "Event Name"),

                         Position =
                        GetInt(rdr, "POS"),

                         DiscountName =
                        GetStr(rdr, "Discount Name")
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

        private static int GetInt(
            SqlDataReader rdr,
            string name)
        {
            if (rdr[name] == DBNull.Value)
                return 0;

            int.TryParse(
                rdr[name]?.ToString(),
                out int value);

            return value;
        }
    }

    public class DiscountEventDto
    {
        public string EventName { get; set; } = "";
        public int Position { get; set; }
        public string DiscountName { get; set; } = "";
    }
}