using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/discount-combos")]
    public class DiscountCombosController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public DiscountCombosController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetDiscountCombos()
        {
            try
            {
                string cs = _configuration.GetConnectionString("DefaultConnection");

                var list = new List<DiscountComboDto>();

                using var cn = new SqlConnection(cs);
                cn.Open();

                string sql = @"
SELECT
   [Match],
   [Discount Name],
   [Discount Amount],
   [Limit],
   [Coupon Type]
FROM [dbo].[DiscountCombos]
ORDER BY [Match],[Discount Name]";

                using var cmd = new SqlCommand(sql, cn);
                using var rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(new DiscountComboDto
                    {
                        MatchValue = GetStr(rdr, "Match"),
                        DiscountName = GetStr(rdr, "Discount Name"),
                        DiscountAmount = GetDecimal(rdr, "Discount Amount"),
                        LimitValue = GetInt(rdr, "Limit"),
                        CouponType = GetStr(rdr, "Coupon Type")
                    });
                }

                return Ok(list);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Error: " + ex.Message);
            }
        }

        private static string GetStr(SqlDataReader rdr, string name)
        {
            return rdr[name] == DBNull.Value
                ? ""
                : rdr[name]?.ToString() ?? "";
        }

        private static int GetInt(SqlDataReader rdr, string name)
        {
            if (rdr[name] == DBNull.Value) return 0;

            int.TryParse(
                rdr[name]?.ToString(),
                out int value);

            return value;
        }

        private static decimal GetDecimal(SqlDataReader rdr, string name)
        {
            if (rdr[name] == DBNull.Value) return 0m;

            decimal.TryParse(
                rdr[name]?.ToString(),
                out decimal value);

            return value;
        }
    }

    public class DiscountComboDto
    {
        public string MatchValue { get; set; } = "";
        public string DiscountName { get; set; } = "";
        public decimal DiscountAmount { get; set; }
        public int LimitValue { get; set; }
        public string CouponType { get; set; } = "";
    }
}