using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/tender-types")]
    public class TenderTypesController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public TenderTypesController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetTenderTypes()
        {
            try
            {
                string cs = _configuration.GetConnectionString("DefaultConnection");

                var list = new List<TenderTypeDto>();

                using var cn = new SqlConnection(cs);
                cn.Open();

                string sql = @"
SELECT
    [Card Name],
    [House Account],
    [Vendor Percentage],
    [Access Level],
    [Is Cash],
    [Is Exchange],
    [Is Bowler],
    [Is Debit],
    [Is House],
    [Is Pin Pad],
    [EMV],
    [TablePay],
    [ConvenienceFeeAmt],
    [Fee Type]
FROM [dbo].[Credit Cards]
ORDER BY [Card Name]";

                using var cmd = new SqlCommand(sql, cn);
                using var rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(new TenderTypeDto
                    {
                        TenderName = GetStr(rdr, "Card Name"),
                        HouseAccount = GetStr(rdr, "House Account"),
                        HotelVendorPercentage = GetDecimal(rdr, "Vendor Percentage"),
                        AccessLevel = GetInt(rdr, "Access Level"),

                        IsCash = ToBool(rdr["Is Cash"]),
                        IsExchange = ToBool(rdr["Is Exchange"]),
                        AvailableOnKiosk = ToBool(rdr["Is Bowler"]),
                        UsesIntercard = ToBool(rdr["Is Debit"]),
                        OpenCashDrawer = ToBool(rdr["Is House"]),
                        UsesGiftEPay = ToBool(rdr["Is Pin Pad"]),
                        SupportsEmv = ToBool(rdr["EMV"]),
                        SupportsTablePay = ToBool(rdr["TablePay"]),

                        ConvenienceFeeAmount = GetDecimal(rdr, "ConvenienceFeeAmt"),
                        FeeType = GetStr(rdr, "Fee Type")
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
            return rdr[name] == DBNull.Value ? "" : rdr[name]?.ToString() ?? "";
        }

        private static int GetInt(SqlDataReader rdr, string name)
        {
            if (rdr[name] == DBNull.Value) return 0;
            int.TryParse(rdr[name]?.ToString(), out int value);
            return value;
        }

        private static decimal GetDecimal(SqlDataReader rdr, string name)
        {
            if (rdr[name] == DBNull.Value) return 0m;
            decimal.TryParse(rdr[name]?.ToString(), out decimal value);
            return value;
        }

        private static bool ToBool(object value)
        {
            if (value == null || value == DBNull.Value) return false;

            string s = value.ToString()?.Trim().ToLower() ?? "";

            return s == "1" || s == "true" || s == "yes" || s == "y";
        }
    }

    public class TenderTypeDto
    {
        public string TenderName { get; set; } = "";
        public string HouseAccount { get; set; } = "";
        public decimal HotelVendorPercentage { get; set; }
        public int AccessLevel { get; set; }

        public bool IsCash { get; set; }
        public bool IsExchange { get; set; }
        public bool AvailableOnKiosk { get; set; }
        public bool UsesIntercard { get; set; }
        public bool OpenCashDrawer { get; set; }
        public bool UsesGiftEPay { get; set; }
        public bool SupportsEmv { get; set; }
        public bool SupportsTablePay { get; set; }

        public decimal ConvenienceFeeAmount { get; set; }
        public string FeeType { get; set; } = "";
    }
}