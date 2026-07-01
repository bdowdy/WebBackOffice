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
        public IActionResult GetClosedChecks(DateTime businessDate, string? searchType,
                                             string? searchValue, string? startTime, string? endTime,
                                             string? mode)
        {
            try
            {
                // Live  = currently open checks (Open Checks, Close Out Day 0, any date)
                // Archive = historical closed checks for the chosen business date
                bool live = string.Equals(mode, "live", StringComparison.OrdinalIgnoreCase);
                string table = live
                    ? "[chk statsql].[dbo].[Open Checks]"
                    : "[chk statsql].[dbo].[Closed Checks]";

                var rows = new List<ClosedCheckListRow>();

                using var cn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                cn.Open();

                // Net sale mirrors the original POS calc:
                // gross sales + taxes - voided sales - (discounts - voided discounts)
                const string netAmount = @"
                    (ISNULL([Gross Sales],0) + ISNULL([Gross Sales Tax],0)
                       + ISNULL([Gross Special Tax],0) + ISNULL([Gross Liquor Tax],0))
                    - ISNULL([Voided Sales],0)
                    - (ABS(ISNULL([Gross Discounts],0)) - ABS(ISNULL([Voided Discounts],0)))";

                // Scope the parent rows and the detail/payment EXISTS subqueries the same way.
                // Live = today's open checks (Close Out Day 0, order date = today).
                string baseFilter = live
                    ? "ISNULL([Close Out Day],0) = 0 AND TRY_CAST([Order Date] AS date) = CAST(GETDATE() AS date)"
                    : "TRY_CAST([Order Date] AS date) = @BusinessDate";

                string detailScope = live
                    ? "ISNULL(d.[Close Out Day],0) = 0"
                    : @"ISNULL(d.[Close Out Day],0) = ISNULL(cc.[Close Out Day],0)
                                       AND TRY_CAST(d.[Check Date] AS date) = @BusinessDate";

                string paymentScope = live
                    ? "ISNULL(p.[Close Out Day],0) = 0"
                    : @"ISNULL(p.[Close Out Day],0) = ISNULL(cc.[Close Out Day],0)
                                       AND TRY_CAST(p.[Date] AS date) = @BusinessDate";

                // Open Checks lacks [Cashier Name] / [Close Time] and stores the open
                // timestamp in [Check Opened] rather than [Open Time].
                string openTimeCol  = live ? "[Check Opened]" : "[Open Time]";
                string cashierCol   = live ? "CAST(NULL AS nvarchar(50))" : "[Cashier Name]";
                string closeTimeCol = live ? "CAST(NULL AS nvarchar(50))" : "[Close Time]";

                string sql = $@"
                    SELECT
                        [Check Number],
                        [Server Name],
                        {cashierCol} AS [Cashier Name],
                        ISNULL([Table Number],'') AS TableNumber,
                        [Order Date],
                        {openTimeCol} AS [Open Time],
                        {closeTimeCol} AS [Close Time],
                        ISNULL([Check Total], 0) AS CheckTotal,
                        ({netAmount}) AS NetAmount,
                        ISNULL([Close Out Day], 0) AS CloseOutDay
                    FROM {table} cc
                    WHERE {baseFilter}";

                switch ((searchType ?? "").ToLowerInvariant())
                {
                    case "check":
                        sql += " AND [Check Number] = @CheckNumber";
                        break;
                    case "server":
                        sql += " AND [Server Name] LIKE @Like";
                        break;
                    case "table":
                        sql += " AND [Table Number] LIKE @Like";
                        break;
                    case "member":
                        sql += " AND [Member Name] LIKE @Like";
                        break;
                    case "voided":
                        sql += " AND ISNULL([Voided Sales],0) > 0";
                        break;
                    case "time":
                        sql += $@" AND TRY_CAST({openTimeCol} AS time) BETWEEN TRY_CAST(@StartTime AS time)
                                                                        AND TRY_CAST(@EndTime AS time)";
                        break;
                    case "menuitem":
                        sql += $@" AND EXISTS (SELECT 1 FROM [chk statsql].[dbo].[Check Detail] d
                                     WHERE d.[Check Number] = cc.[Check Number]
                                       AND d.[Menu Item] LIKE @Like
                                       AND {detailScope})";
                        break;
                    case "payment":
                        sql += $@" AND EXISTS (SELECT 1 FROM [chk statsql].[dbo].[Payments] p
                                     WHERE p.[Check Number] = cc.[Check Number]
                                       AND p.[Payment Type] LIKE @Like
                                       AND {paymentScope})";
                        break;
                }

                sql += $" ORDER BY TRY_CAST({openTimeCol} AS datetime), [Check Number]";

                using var cmd = new SqlCommand(sql, cn);

                cmd.Parameters.AddWithValue("@BusinessDate", businessDate.Date);
                cmd.Parameters.AddWithValue("@Like", "%" + (searchValue ?? "") + "%");
                cmd.Parameters.AddWithValue("@CheckNumber", int.TryParse(searchValue, out var ck) ? ck : 0);
                cmd.Parameters.AddWithValue("@StartTime", (object?)(string.IsNullOrWhiteSpace(startTime) ? "00:00" : startTime));
                cmd.Parameters.AddWithValue("@EndTime", (object?)(string.IsNullOrWhiteSpace(endTime) ? "23:59" : endTime));

                using var rd = cmd.ExecuteReader();

                while (rd.Read())
                {
                    rows.Add(new ClosedCheckListRow
                    {
                        CheckNumber = Convert.ToInt32(rd["Check Number"]),
                        ServerName = rd["Server Name"].ToString() ?? "",
                        CashierName = rd["Cashier Name"].ToString() ?? "",
                        TableNumber = rd["TableNumber"].ToString() ?? "",
                        OrderDate = FormatDate(rd["Order Date"]),
                        OpenTime = FormatTime(rd["Open Time"]),
                        CloseTime = FormatTime(rd["Close Time"]),
                        CheckTotal = Convert.ToDecimal(rd["CheckTotal"]),
                        NetAmount = Convert.ToDecimal(rd["NetAmount"]),
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
        public IActionResult GetCheckReview(int checkNumber, DateTime businessDate, int closeOutDay, string? mode)
        {
            try
            {
                bool live = string.Equals(mode, "live", StringComparison.OrdinalIgnoreCase);
                string table = live
                    ? "[chk statsql].[dbo].[Open Checks]"
                    : "[chk statsql].[dbo].[Closed Checks]";

                var result = new CheckReviewResult();
                var details = new List<CheckDetailRow>();
                var payments = new List<CheckPaymentRow>();

                using var cn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                cn.Open();

                string headerSql = live
                    ? $@"
                    SELECT TOP 1 *
                    FROM {table}
                    WHERE [Check Number] = @CheckNumber
                      AND ISNULL([Close Out Day], 0) = 0"
                    : $@"
                    SELECT TOP 1 *
                    FROM {table}
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
                        cashierName = SafeStr(rd, "Cashier Name"),
                        tableNumber = rd["Table Number"].ToString() ?? "",
                        orderType = SafeStr(rd, "Order Destination"),
                        partyNumber = rd["Number In Party"].ToString() ?? "",
                        orderDate = FormatDate(rd["Order Date"]),
                        openTime = FormatTimeSafe(rd, live ? "Check Opened" : "Open Time"),
                        closeTime = FormatTimeSafe(rd, "Close Time"),
                        checkTotal = SafeDecimal(rd, "Check Total"),
                        grossSales = SafeDecimal(rd, "Gross Sales"),
                        salesTax = SafeDecimal(rd, "Gross Sales Tax"),
                        specialTax = SafeDecimal(rd, "Gross Special Tax"),
                        discounts = SafeDecimal(rd, "Gross Discounts"),
                        voidedSales = SafeDecimal(rd, "Voided Sales"),
                        memberName = rd["Member Name"].ToString() ?? ""
                    };
                }

                string detailSql = live
                    ? @"
                    SELECT *
                    FROM [chk statsql].[dbo].[Check Detail]
                    WHERE [Check Number] = @CheckNumber
                      AND ISNULL([Close Out Day], 0) = 0
                    ORDER BY ISNULL([Seat Number], 0), ABS(ISNULL([Selection Number], 0)), ISNULL([Pos], 0)"
                    : @"
                    SELECT *
                    FROM [chk statsql].[dbo].[Check Detail]
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
                            Amount = SafeDecimal(rd, "Price"),
                            Voided = SafeDecimal(rd, "Voided") != 0,
                            VoidReason = HasColumn(rd, "Void Reason") ? rd["Void Reason"].ToString() ?? "" : "",
                            VoidedBy = HasColumn(rd, "Voided By") ? rd["Voided By"].ToString() ?? "" : "",
                            VoidTime = HasColumn(rd, "Void Time") ? FormatTime(rd["Void Time"]) : ""
                        });
                    }
                }

                string paymentSql = live
                    ? @"
                    SELECT *
                    FROM [chk statsql].[dbo].[Payments]
                    WHERE [Check Number] = @CheckNumber
                      AND ISNULL([Close Out Day], 0) = 0
                    ORDER BY TRY_CAST([Time] AS datetime)"
                    : @"
                    SELECT *
                    FROM [chk statsql].[dbo].[Payments]
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

        private static string SafeStr(SqlDataReader rd, string column)
        {
            return HasColumn(rd, column) && rd[column] != DBNull.Value
                ? rd[column].ToString() ?? ""
                : "";
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

        private static string FormatTimeSafe(SqlDataReader rd, string column)
        {
            return HasColumn(rd, column) ? FormatTime(rd[column]) : "";
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
            public string TableNumber { get; set; } = "";
            public string OrderDate { get; set; } = "";
            public string OpenTime { get; set; } = "";
            public string CloseTime { get; set; } = "";
            public decimal CheckTotal { get; set; }
            public decimal NetAmount { get; set; }
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