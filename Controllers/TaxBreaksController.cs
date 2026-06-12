using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/tax-breaks")]
    public class TaxBreaksController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public TaxBreaksController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetTaxBreaks()
        {
            try
            {
                string cs = _configuration.GetConnectionString("DefaultConnection");
                var list = new List<TaxBreakDto>();

                using var cn = new SqlConnection(cs);
                cn.Open();

                string sql = @"
SELECT
    [Tax Table],
    [Start],
    [End],
    [Tax],
    [Row]
FROM [dbo].[Tax Breaks]
ORDER BY [Tax Table], [Row]";

                using var cmd = new SqlCommand(sql, cn);
                using var rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(new TaxBreakDto
                    {
                        TaxTable = rdr["Tax Table"]?.ToString() ?? "",
                        Start = rdr["Start"] == DBNull.Value ? 0m : Convert.ToDecimal(rdr["Start"]),
                        End = rdr["End"] == DBNull.Value ? 0m : Convert.ToDecimal(rdr["End"]),
                        Tax = rdr["Tax"] == DBNull.Value ? 0m : Convert.ToDecimal(rdr["Tax"]),
                        Row = rdr["Row"] == DBNull.Value ? 0 : Convert.ToInt32(rdr["Row"])
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

    public class TaxBreakDto
    {
        public string TaxTable { get; set; } = "";
        public decimal Start { get; set; }
        public decimal End { get; set; }
        public decimal Tax { get; set; }
        public int Row { get; set; }
    }
}