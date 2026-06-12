using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/coupon-assignments")]
    public class CouponAssignmentsController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public CouponAssignmentsController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetCouponAssignments([FromQuery] string couponName)
        {
            var list = new List<string>();

            if (string.IsNullOrWhiteSpace(couponName))
                return Ok(list);

            //string connStr = _configuration.GetConnectionString("MSaleData");
            string cs = _configuration.GetConnectionString("DefaultConnection");

            using var cn = new SqlConnection(cs);
            {
                cn.Open();

                string sql = @"
                    SELECT [LinkToName]
                    FROM [CouponAssignments]
                    WHERE [CouponName] = @CouponName
                    ORDER BY [LinkToName]";

                using (SqlCommand cmd = new SqlCommand(sql, cn))
                {
                    cmd.Parameters.AddWithValue("@CouponName", couponName);

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            list.Add(dr["LinkToName"].ToString().Trim());
                        }
                    }
                }
            }

            return Ok(list);
        }

        public class CouponAssignmentDto
        {
            public string CouponName { get; set; } = "";
            public string LinkToName { get; set; } = "";
        }
    }
}