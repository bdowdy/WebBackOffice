using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/coupon-reasons")]
    public class CouponReasonsController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public CouponReasonsController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetCouponReasons()
        {
            try
            {
                string cs = _configuration.GetConnectionString("DefaultConnection");

                var list = new List<CouponReasonDto>();

                using var cn = new SqlConnection(cs);
                cn.Open();

                string sql = @"
SELECT [Name]
FROM [dbo].[Coupon Reason]
ORDER BY [Name]";

                using var cmd = new SqlCommand(sql, cn);
                using var rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(new CouponReasonDto
                    {
                        Name = rdr["Name"]?.ToString() ?? ""
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

    public class CouponReasonDto
    {
        public string Name { get; set; } = "";
    }
}