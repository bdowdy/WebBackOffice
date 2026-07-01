using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    public class TimeRecordsReportController : Controller
    {
        private readonly IConfiguration _configuration;

        public TimeRecordsReportController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View("~/Views/MSaleAdmin/TimeRecordsReport.cshtml");
        }

        // GET /TimeRecordsReport/GetTimeRecords?beginDate=2026-06-01&endDate=2026-06-07
        [HttpGet]
        public IActionResult GetTimeRecords(DateTime beginDate, DateTime endDate)
        {
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();

            // Mirror of PrintPayRoll / SyncPayrollData query
            // Groups records by employee, calculates reg/OT hours (40hr weekly threshold)
            string sql = @"
                SELECT
                    [Employee Name]                          AS EmployeeName,
                    [Department Name]                        AS DepartmentName,
                    [Shift Date]                             AS ShiftDate,
                    [Punch Out Date]                         AS PunchOutDate,
                    [Time In]                                AS TimeIn,
                    [Time Out]                               AS TimeOut,
                    [Break Out]                              AS BreakOut,
                    [Break In]                               AS BreakIn,
                    ISNULL([Rate of Pay],       0)           AS RateOfPay,
                    ISNULL([Tip Amount],        0)           AS DeclaredTips,
                    ISNULL([Credit Card Tips],  0)           AS CreditCardTips,
                    ISNULL([Tips Paid Out],     0)           AS TipsPaidOut,
                    ISNULL([Tip Share],         0)           AS TipShare,
                    ISNULL([Sales],             0)           AS Sales,
                    ISNULL([Shift],             0)           AS Shift,
                    ISNULL([Paid By],           1)           AS PaidBy,
                    ISNULL([Altered By],        '')          AS AlteredBy
                FROM [MSaleData].[dbo].[Employee Records]
                WHERE TRY_CAST([Shift Date] AS datetime) BETWEEN @BeginDate AND @EndDate
                ORDER BY [Employee Name],
                         TRY_CAST([Shift Date] AS datetime),
                         TRY_CAST([Time In]    AS datetime)";

            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@BeginDate", beginDate.Date);
            cmd.Parameters.AddWithValue("@EndDate",   endDate.Date.AddDays(1).AddSeconds(-1));

            using var rd = cmd.ExecuteReader();

            var rows   = new List<TimeRecordRow>();
            var empHrs = new Dictionary<string, decimal>();   // running weekly hours per employee

            while (rd.Read())
            {
                string emp  = rd["EmployeeName"].ToString()  ?? "";
                string dept = rd["DepartmentName"].ToString() ?? "";

                DateTime shiftDate    = TryDate(rd["ShiftDate"],   DateTime.Today);
                DateTime punchOutDate = TryDate(rd["PunchOutDate"], shiftDate);

                DateTime? tIn   = BuildDateTime(shiftDate,    rd["TimeIn"]);
                DateTime? tOut  = BuildDateTime(punchOutDate, rd["TimeOut"]);
                DateTime? bkOut = BuildDateTime(shiftDate,    rd["BreakOut"]);
                DateTime? bkIn  = BuildDateTime(shiftDate,    rd["BreakIn"]);

                decimal worked = 0;
                if (tIn.HasValue && tOut.HasValue && tOut > tIn)
                    worked = (decimal)(tOut.Value - tIn.Value).TotalHours;

                if (bkOut.HasValue && bkIn.HasValue && bkIn > bkOut)
                    worked -= (decimal)(bkIn.Value - bkOut.Value).TotalHours;

                worked = Math.Max(0, Math.Round(worked, 2));

                if (!empHrs.ContainsKey(emp)) empHrs[emp] = 0;
                decimal before = empHrs[emp];
                decimal after  = before + worked;

                decimal reg = 0, ot = 0;
                if (before >= 40)         { ot  = worked; }
                else if (after > 40)      { reg = 40 - before; ot = after - 40; }
                else                      { reg = worked; }

                empHrs[emp] = after;

                decimal rate  = Convert.ToDecimal(rd["RateOfPay"]);
                decimal wages = (reg * rate) + (ot * rate * 1.5m);

                rows.Add(new TimeRecordRow
                {
                    EmployeeName   = emp,
                    DepartmentName = dept,
                    ShiftDate      = shiftDate.ToString("MM/dd/yyyy"),
                    TimeIn         = FormatTime(rd["TimeIn"]),
                    TimeOut        = FormatTime(rd["TimeOut"]),
                    BreakOut       = FormatTime(rd["BreakOut"]),
                    BreakIn        = FormatTime(rd["BreakIn"]),
                    RegHours       = Math.Round(reg,   2),
                    OTHours        = Math.Round(ot,    2),
                    TotalHours     = worked,
                    RateOfPay      = rate,
                    Wages          = Math.Round(wages, 2),
                    DeclaredTips   = Convert.ToDecimal(rd["DeclaredTips"]),
                    CreditCardTips = Convert.ToDecimal(rd["CreditCardTips"]),
                    TipsPaidOut    = Convert.ToDecimal(rd["TipsPaidOut"]),
                    TipShare       = Convert.ToDecimal(rd["TipShare"]),
                    Sales          = Convert.ToDecimal(rd["Sales"]),
                    AlteredBy      = rd["AlteredBy"].ToString() ?? ""
                });
            }

            return Json(rows);
        }

        // GET /TimeRecordsReport/GetEmployeeList?beginDate=...&endDate=...
        [HttpGet]
        public IActionResult GetEmployeeList(DateTime beginDate, DateTime endDate)
        {
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            var names = cn.Query<string>(@"
                SELECT DISTINCT [Employee Name]
                FROM [MSaleData].[dbo].[Employee Records]
                WHERE TRY_CAST([Shift Date] AS datetime) BETWEEN @BeginDate AND @EndDate
                ORDER BY [Employee Name]",
                new { BeginDate = beginDate.Date, EndDate = endDate.Date.AddDays(1).AddSeconds(-1) });
            return Json(names);
        }

        private static DateTime TryDate(object val, DateTime fallback)
        {
            if (val == DBNull.Value || val == null) return fallback;
            return DateTime.TryParse(val.ToString(), out var d) ? d : fallback;
        }

        private static DateTime? BuildDateTime(DateTime baseDate, object val)
        {
            if (val == DBNull.Value || val == null) return null;
            if (!DateTime.TryParse(val.ToString(), out var t)) return null;
            return baseDate.Date.Add(t.TimeOfDay);
        }

        private static string FormatTime(object val)
        {
            if (val == DBNull.Value || val == null) return "";
            if (!DateTime.TryParse(val.ToString(), out var t)) return "";
            return t.ToString("hh:mm tt");
        }
    }

    public class TimeRecordRow
    {
        public string  EmployeeName   { get; set; } = "";
        public string  DepartmentName { get; set; } = "";
        public string  ShiftDate      { get; set; } = "";
        public string  TimeIn         { get; set; } = "";
        public string  TimeOut        { get; set; } = "";
        public string  BreakOut       { get; set; } = "";
        public string  BreakIn        { get; set; } = "";
        public decimal RegHours       { get; set; }
        public decimal OTHours        { get; set; }
        public decimal TotalHours     { get; set; }
        public decimal RateOfPay      { get; set; }
        public decimal Wages          { get; set; }
        public decimal DeclaredTips   { get; set; }
        public decimal CreditCardTips { get; set; }
        public decimal TipsPaidOut    { get; set; }
        public decimal TipShare       { get; set; }
        public decimal Sales          { get; set; }
        public string  AlteredBy      { get; set; } = "";
    }
}
