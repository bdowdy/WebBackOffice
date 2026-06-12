using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/special-tax")]
    public class SpecialTaxController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public SpecialTaxController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetSpecialTax()
        {
            var list = new List<SpecialTaxDto>();
            string cs = _configuration.GetConnectionString("DefaultConnection");

            using var cn = new SqlConnection(cs);
            cn.Open();

            string sql = @"
SELECT [Special Tax], [Min Amount], [Max Amount], [Exemption], [Tax Name], [TaxTable]
FROM [dbo].[Special Tax Info]
ORDER BY [Tax Name], [Special Tax]";

            using var cmd = new SqlCommand(sql, cn);
            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                list.Add(new SpecialTaxDto
                {
                    SpecialTax = rdr["Special Tax"]?.ToString() ?? "",
                    MinAmount = rdr["Min Amount"] == DBNull.Value ? 0 : Convert.ToDecimal(rdr["Min Amount"]),
                    MaxAmount = rdr["Max Amount"] == DBNull.Value ? 0 : Convert.ToDecimal(rdr["Max Amount"]),
                    Exemption = rdr["Exemption"]?.ToString() ?? "",
                    TaxName = rdr["Tax Name"]?.ToString() ?? "",
                    TaxTable = rdr["TaxTable"]?.ToString() ?? ""
                });
            }

            return Ok(list);
        }
    }

    public class SpecialTaxDto
    {
        public string SpecialTax { get; set; } = "";
        public decimal MinAmount { get; set; }
        public decimal MaxAmount { get; set; }
        public string Exemption { get; set; } = "";
        public string TaxName { get; set; } = "";
        public string TaxTable { get; set; } = "";
    }
}