using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/sales-tax")]
    public class SalesTaxController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public SalesTaxController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetSalesTax()
        {
            try
            {
                string cs = _configuration.GetConnectionString("DefaultConnection");
                var list = new List<SalesTaxDto>();

                using var cn = new SqlConnection(cs);
                cn.Open();

                string sql = @"
SELECT
    [Tax Name],
    [Tax Amount],
    [TaxTable]
FROM [dbo].[Sales Tax]
ORDER BY [TaxTable], [Tax Name]";

                using var cmd = new SqlCommand(sql, cn);
                using var rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(new SalesTaxDto
                    {
                        TaxName = rdr["Tax Name"]?.ToString() ?? "",
                        TaxAmount = rdr["Tax Amount"] == DBNull.Value ? 0m : Convert.ToDecimal(rdr["Tax Amount"]),
                        TaxTable = rdr["TaxTable"]?.ToString() ?? ""
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

    public class SalesTaxDto
    {
        public string TaxName { get; set; } = "";
        public decimal TaxAmount { get; set; }
        public string TaxTable { get; set; } = "";
    }
}