using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    // Web port of the legacy VB "GetCreditCardReport" routine.
    // Server tip / sales income from [Tip Income], grouped by server over a date range.
    public class ServerIncomeReportController : Controller
    {
        private readonly IConfiguration _cfg;
        public ServerIncomeReportController(IConfiguration cfg) => _cfg = cfg;

        private SqlConnection Cn() => new(_cfg.GetConnectionString("DefaultConnection"));

        [HttpGet]
        public IActionResult Index() =>
            View("~/Views/MSaleAdmin/ServerIncomeReport.cshtml");

        [HttpGet]
        public IActionResult GetServerIncome(DateTime beginDate, DateTime endDate)
        {
            try
            {
                using var cn = Cn();
                cn.Open();

                const string sql = @"
                    SELECT
                        [Server Name]                       AS ServerName,
                        SUM(ISNULL([Tip Sales],0))          AS TipSales,
                        SUM(ISNULL([Declared Tips],0))      AS DeclaredTips,
                        SUM(ISNULL([Charged Tips],0))       AS ChargedTips,
                        SUM(ISNULL([Gratuity],0))           AS Gratuity,
                        SUM(ISNULL([Tips Paid Out],0))      AS TipsPaidOut,
                        SUM(ISNULL([Charge Sales],0))       AS ChargeSales,
                        SUM(ISNULL([Tip Share],0))          AS TipShare
                    FROM [MSaleData].[dbo].[Tip Income]
                    WHERE TRY_CAST([Date] AS date) BETWEEN @Begin AND @End
                      AND ISNULL([Server Name],'') <> ''
                    GROUP BY [Server Name]
                    ORDER BY [Server Name]";

                using var cmd = new SqlCommand(sql, cn);
                cmd.Parameters.AddWithValue("@Begin", beginDate.Date);
                cmd.Parameters.AddWithValue("@End", endDate.Date);

                using var rd = cmd.ExecuteReader();

                var rows = new List<object>();
                decimal tTipSales = 0, tChgSales = 0, tCashTips = 0, tChgTips = 0,
                        tPaidOut = 0, tShare = 0, tTotalTips = 0;

                while (rd.Read())
                {
                    decimal tipSales = D(rd["TipSales"]);
                    decimal chgSales = D(rd["ChargeSales"]);
                    decimal declared = D(rd["DeclaredTips"]);
                    decimal charged = D(rd["ChargedTips"]);
                    decimal gratuity = D(rd["Gratuity"]);
                    decimal paidOut = D(rd["TipsPaidOut"]);
                    decimal tipShare = D(rd["TipShare"]);

                    decimal cashTips = declared;             // "Cash Tips" = declared tips
                    decimal chgTips = charged + gratuity;    // "Chg Tips"  = charged + gratuity
                    decimal totalTips = declared + tipShare + charged + gratuity - paidOut;
                    decimal? chgTipPct = chgSales > 0 && chgTips > 0
                        ? Math.Round(chgTips / chgSales * 100, 2)
                        : (decimal?)null;

                    rows.Add(new
                    {
                        serverName = rd["ServerName"]?.ToString() ?? "",
                        tipSales,
                        chgSales,
                        cashTips,
                        chgTips,
                        tipsPaidOut = paidOut,
                        tipShare,
                        totalTips,
                        chgTipPct
                    });

                    tTipSales += tipSales;
                    tChgSales += chgSales;
                    tCashTips += cashTips;
                    tChgTips += chgTips;
                    tPaidOut += paidOut;
                    tShare += tipShare;
                    tTotalTips += totalTips;
                }

                var totals = new
                {
                    tipSales = tTipSales,
                    chgSales = tChgSales,
                    cashTips = tCashTips,
                    chgTips = tChgTips,
                    tipsPaidOut = tPaidOut,
                    tipShare = tShare,
                    totalTips = tTotalTips,
                    chgTipPct = tChgSales > 0 && tChgTips > 0
                        ? Math.Round(tChgTips / tChgSales * 100, 2)
                        : (decimal?)null
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
    }
}
