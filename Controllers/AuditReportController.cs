using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    // Web port of the legacy VB "ShowAllAudits" routine.
    // Reads [AuditSales] for a date range and returns per-register/operator
    // close-out rows plus column totals.
    public class AuditReportController : Controller
    {
        private readonly IConfiguration _cfg;
        public AuditReportController(IConfiguration cfg) => _cfg = cfg;

        private SqlConnection Cn() => new(_cfg.GetConnectionString("DefaultConnection"));

        [HttpGet]
        public IActionResult Index() =>
            View("~/Views/MSaleAdmin/AuditReport.cshtml");

        [HttpGet]
        public IActionResult GetAuditSales(DateTime beginDate, DateTime endDate)
        {
            try
            {
                using var cn = Cn();
                cn.Open();

                // Skim can be stored as text/blank in the legacy schema, so coerce it.
                const string sql = @"
                    SELECT
                        [Register Name]                          AS RegisterName,
                        [Closed By]                              AS Manager,
                        [Date]                                   AS AuditDate,
                        [Start Time]                             AS StartTime,
                        [End Time]                               AS EndTime,
                        [Operator]                               AS OperatorName,
                        ISNULL([Total To Account For],0)         AS Sales,
                        ISNULL([Discounts],0)                    AS Discounts,
                        ISNULL([Voids],0)                        AS Voids,
                        ISNULL([Cash In Drawer],0)               AS CashInDrawer,
                        ISNULL([Bank Deposit],0)                 AS BankDeposit,
                        ISNULL(TRY_CONVERT(money,[Skim]),0)      AS Skim,
                        ISNULL([Tips],0)                         AS Tips
                    FROM [MSaleData].[dbo].[AuditSales]
                    WHERE TRY_CAST([Date] AS date) BETWEEN @Begin AND @End
                    ORDER BY TRY_CAST([Date] AS datetime), TRY_CAST([Start Time] AS datetime)";

                using var cmd = new SqlCommand(sql, cn);
                cmd.Parameters.AddWithValue("@Begin", beginDate.Date);
                cmd.Parameters.AddWithValue("@End", endDate.Date);

                using var rd = cmd.ExecuteReader();

                var rows = new List<object>();
                decimal tSales = 0, tDisc = 0, tVoid = 0, tCash = 0, tDep = 0, tOver = 0, tTip = 0;

                while (rd.Read())
                {
                    decimal sales = D(rd["Sales"]);
                    decimal disc = D(rd["Discounts"]);
                    decimal voids = D(rd["Voids"]);
                    decimal cashInDrawer = D(rd["CashInDrawer"]);
                    decimal cashDeposit = D(rd["BankDeposit"]) + D(rd["Skim"]);
                    decimal tips = D(rd["Tips"]);

                    // Legacy rule: show 0 for +/- when the drawer is negative.
                    decimal overShort = cashInDrawer < 0 ? 0 : cashDeposit - cashInDrawer;

                    rows.Add(new
                    {
                        registerName = rd["RegisterName"]?.ToString() ?? "",
                        manager = rd["Manager"]?.ToString() ?? "",
                        auditDate = FmtDate(rd["AuditDate"]),
                        startTime = FmtTime(rd["StartTime"]),
                        endTime = FmtTime(rd["EndTime"]),
                        operatorName = rd["OperatorName"]?.ToString() ?? "",
                        sales,
                        discounts = disc,
                        voids,
                        cashInDrawer,
                        cashDeposit,
                        overShort,
                        tips
                    });

                    tSales += sales;
                    tDisc += disc;
                    tVoid += voids;
                    tCash += cashInDrawer;
                    tDep += cashDeposit;
                    tOver += overShort;
                    tTip += tips;
                }

                var totals = new
                {
                    sales = tSales,
                    discounts = tDisc,
                    voids = tVoid,
                    cashInDrawer = tCash,
                    cashDeposit = tDep,
                    overShort = tOver,
                    tips = tTip
                };

                return Json(new { success = true, data = rows, totals });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        private static decimal D(object v) =>
            v == null || v == DBNull.Value ? 0
            : decimal.TryParse(v.ToString(), out var d) ? d : 0;

        private static string FmtDate(object v) =>
            v == DBNull.Value || v == null
                ? ""
                : DateTime.TryParse(v.ToString(), out var d) ? d.ToString("MM/dd/yy") : v.ToString() ?? "";

        private static string FmtTime(object v) =>
            v == DBNull.Value || v == null
                ? ""
                : DateTime.TryParse(v.ToString(), out var d) ? d.ToString("h:mm tt") : v.ToString() ?? "";
    }
}
