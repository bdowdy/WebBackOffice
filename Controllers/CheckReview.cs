using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    public class CheckReviewController : Controller
    {
        private readonly IConfiguration _configuration;

        public CheckReviewController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View("~/Views/MSaleAdmin/CheckReview.cshtml");
        }

        [HttpGet]
        public IActionResult GetClosedChecks(DateTime businessDate, string? searchType, string? searchValue)
        {
            try
            {
                var rows = new List<ClosedCheckListRow>();

                using var cn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                cn.Open();

                string sql = @"
                    SELECT
                        [Check Number],
                        [Server Name],
                        [Cashier Name],
                        [Order Date],
                        [Open Time],
                        [Close Time],
                        ISNULL([Check Total], 0) AS CheckTotal,
                        ISNULL([Close Out Day], 0) AS CloseOutDay
                    FROM [dbo].[Closed Checks]
                    WHERE TRY_CAST([Order Date] AS date) = @BusinessDate
                ";

                if (searchType == "server")
                    sql += " AND [Server Name] LIKE @SearchValue ";

                if (searchType == "check")
                    sql += " AND [Check Number] = @CheckNumber ";

                sql += " ORDER BY TRY_CAST([Open Time] AS datetime), [Check Number]";

                using var cmd = new SqlCommand(sql, cn);

                cmd.Parameters.AddWithValue("@BusinessDate", businessDate.Date);
                cmd.Parameters.AddWithValue("@SearchValue", "%" + (searchValue ?? "") + "%");
                cmd.Parameters.AddWithValue("@CheckNumber", int.TryParse(searchValue, out var ck) ? ck : 0);

                using var rd = cmd.ExecuteReader();

                while (rd.Read())
                {
                    rows.Add(new ClosedCheckListRow
                    {
                        CheckNumber = Convert.ToInt32(rd["Check Number"]),
                        ServerName = rd["Server Name"].ToString() ?? "",
                        CashierName = rd["Cashier Name"].ToString() ?? "",
                        OrderDate = FormatDate(rd["Order Date"]),
                        OpenTime = FormatTime(rd["Open Time"]),
                        CloseTime = FormatTime(rd["Close Time"]),
                        CheckTotal = Convert.ToDecimal(rd["CheckTotal"]),
                        CloseOutDay = Convert.ToInt32(rd["CloseOutDay"])
                    });
                }

                return Json(new { success = true, data = rows });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public IActionResult GetCheckReview(int checkNumber, DateTime businessDate, int closeOutDay)
        {
            try
            {
                var result = new CheckReviewResult();
                var details = new List<CheckDetailRow>();
                var payments = new List<CheckPaymentRow>();

                using var cn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                cn.Open();

                string headerSql = @"
                    SELECT TOP 1 *
                    FROM [dbo].[Closed Checks]
                    WHERE [Check Number] = @CheckNumber
                      AND TRY_CAST([Order Date] AS date) = @BusinessDate
                      AND ISNULL([Close Out Day], 0) = @CloseOutDay";

                using (var cmd = new SqlCommand(headerSql, cn))
                {
                    cmd.Parameters.AddWithValue("@CheckNumber", checkNumber);
                    cmd.Parameters.AddWithValue("@BusinessDate", businessDate.Date);
                    cmd.Parameters.AddWithValue("@CloseOutDay", closeOutDay);

                    using var rd = cmd.ExecuteReader();

                    if (!rd.Read())
                        return Json(new { success = false, message = "Check not found." });

                    result.Header = new
                    {
                        checkNumber,
                        serverName = rd["Server Name"].ToString() ?? "",
                        cashierName = rd["Cashier Name"].ToString() ?? "",
                        tableNumber = rd["Table Number"].ToString() ?? "",
                        orderType = rd["Order Type"].ToString() ?? "",
                        partyNumber = rd["Number In Party"].ToString() ?? "",
                        orderDate = FormatDate(rd["Order Date"]),
                        openTime = FormatTime(rd["Open Time"]),
                        closeTime = FormatTime(rd["Close Time"]),
                        checkTotal = SafeDecimal(rd, "Check Total"),
                        grossSales = SafeDecimal(rd, "Gross Sales"),
                        salesTax = SafeDecimal(rd, "Gross Sales Tax"),
                        specialTax = SafeDecimal(rd, "Gross Special Tax"),
                        discounts = SafeDecimal(rd, "Gross Discounts"),
                        voidedSales = SafeDecimal(rd, "Voided Sales"),
                        memberName = rd["Member Name"].ToString() ?? ""
                    };
                }

                string detailSql = @"
                    SELECT *
                    FROM [dbo].[Check Detail]
                    WHERE [Check Number] = @CheckNumber
                      AND ISNULL([Close Out Day], 0) = @CloseOutDay
                      AND TRY_CAST([Check Date] AS date) = @BusinessDate
                    ORDER BY ISNULL([Seat Number], 0), ABS(ISNULL([Selection Number], 0)), ISNULL([Pos], 0)";

                using (var cmd = new SqlCommand(detailSql, cn))
                {
                    cmd.Parameters.AddWithValue("@CheckNumber", checkNumber);
                    cmd.Parameters.AddWithValue("@BusinessDate", businessDate.Date);
                    cmd.Parameters.AddWithValue("@CloseOutDay", closeOutDay);

                    using var rd = cmd.ExecuteReader();

                    while (rd.Read())
                    {
                        details.Add(new CheckDetailRow
                        {
                            SeatNumber = Convert.ToInt32(SafeDecimal(rd, "Seat Number")),
                            MenuItem = rd["Menu Item"].ToString() ?? "",
                            Amount = SafeDecimal(rd, "Amount"),
                            Voided = SafeDecimal(rd, "Voided") != 0,
                            VoidReason = HasColumn(rd, "Void Reason") ? rd["Void Reason"].ToString() ?? "" : "",
                            VoidedBy = HasColumn(rd, "Voided By") ? rd["Voided By"].ToString() ?? "" : "",
                            VoidTime = HasColumn(rd, "Void Time") ? FormatTime(rd["Void Time"]) : ""
                        });
                    }
                }

                string paymentSql = @"
                    SELECT *
                    FROM [dbo].[Payments]
                    WHERE [Check Number] = @CheckNumber
                      AND ISNULL([Close Out Day], 0) = @CloseOutDay
                      AND TRY_CAST([Date] AS date) = @BusinessDate
                    ORDER BY TRY_CAST([Time] AS datetime)";

                using (var cmd = new SqlCommand(paymentSql, cn))
                {
                    cmd.Parameters.AddWithValue("@CheckNumber", checkNumber);
                    cmd.Parameters.AddWithValue("@BusinessDate", businessDate.Date);
                    cmd.Parameters.AddWithValue("@CloseOutDay", closeOutDay);

                    using var rd = cmd.ExecuteReader();

                    while (rd.Read())
                    {
                        payments.Add(new CheckPaymentRow
                        {
                            PaymentType = rd["Payment Type"].ToString() ?? "",
                            PaymentAmount = SafeDecimal(rd, "Payment Amount"),
                            TipAmount = SafeDecimal(rd, "Tip Amount")
                        });
                    }
                }

                result.Details = details;
                result.Payments = payments;
                result.Totals = new
                {
                    itemTotal = details.Where(x => !x.Voided).Sum(x => x.Amount),
                    voidTotal = details.Where(x => x.Voided).Sum(x => x.Amount),
                    paymentTotal = payments.Sum(x => x.PaymentAmount),
                    tipTotal = payments.Sum(x => x.TipAmount)
                };

                return Json(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        private static decimal SafeDecimal(SqlDataReader rd, string column)
        {
            try
            {
                if (!HasColumn(rd, column) || rd[column] == DBNull.Value)
                    return 0;

                return Convert.ToDecimal(rd[column]);
            }
            catch
            {
                return 0;
            }
        }

        private static bool HasColumn(SqlDataReader rd, string columnName)
        {
            for (int i = 0; i < rd.FieldCount; i++)
            {
                if (string.Equals(rd.GetName(i), columnName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        private static string FormatDate(object value)
        {
            if (value == DBNull.Value || value == null)
                return "";

            return DateTime.TryParse(value.ToString(), out var d)
                ? d.ToString("yyyy-MM-dd")
                : "";
        }

        private static string FormatTime(object value)
        {
            if (value == DBNull.Value || value == null)
                return "";

            return DateTime.TryParse(value.ToString(), out var d)
                ? d.ToString("HH:mm")
                : "";
        }

        public class ClosedCheckListRow
        {
            public int CheckNumber { get; set; }
            public string ServerName { get; set; } = "";
            public string CashierName { get; set; } = "";
            public string OrderDate { get; set; } = "";
            public string OpenTime { get; set; } = "";
            public string CloseTime { get; set; } = "";
            public decimal CheckTotal { get; set; }
            public int CloseOutDay { get; set; }
        }

        public class CheckReviewResult
        {
            public object? Header { get; set; }
            public List<CheckDetailRow> Details { get; set; } = new();
            public List<CheckPaymentRow> Payments { get; set; } = new();
            public object? Totals { get; set; }
        }

        public class CheckDetailRow
        {
            public int SeatNumber { get; set; }
            public string MenuItem { get; set; } = "";
            public decimal Amount { get; set; }
            public bool Voided { get; set; }
            public string VoidReason { get; set; } = "";
            public string VoidedBy { get; set; } = "";
            public string VoidTime { get; set; } = "";
        }

        public class CheckPaymentRow
        {
            public string PaymentType { get; set; } = "";
            public decimal PaymentAmount { get; set; }
            public decimal TipAmount { get; set; }
        }
    }
}