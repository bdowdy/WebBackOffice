using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    public class TimeReportsController : Controller
    {
        private readonly IConfiguration _cfg;
        public TimeReportsController(IConfiguration cfg) => _cfg = cfg;

        SqlConnection Cn() => new(_cfg.GetConnectionString("DefaultConnection"));

        [HttpGet] public IActionResult Index() =>
            View("~/Views/MSaleAdmin/TimeReports.cshtml");

        // ── 1. Pay Period Labor ──────────────────────────────────────────────
        // HouseLabor BeginWeek, EndingDate
        // Hours and wages per employee per day in the period
        [HttpGet]
        public IActionResult GetPayPeriodLabor(DateTime beginDate, DateTime endDate)
        {
            using var cn = Cn(); cn.Open();
            var sql = @"
                SELECT [Employee Name] AS EmployeeName,
                       [Department Name] AS DepartmentName,
                       [Shift Date] AS ShiftDate,
                       [Time In] AS TimeIn, [Time Out] AS TimeOut,
                       [Break Out] AS BreakOut, [Break In] AS BreakIn,
                       ISNULL([Rate of Pay],0) AS RateOfPay,
                       ISNULL([Paid By],1) AS PaidBy,
                       ISNULL([Shift],0) AS Shift
                FROM [MSaleData].[dbo].[Employee Records]
                WHERE TRY_CAST([Shift Date] AS datetime) BETWEEN @B AND @E
                ORDER BY [Employee Name], TRY_CAST([Shift Date] AS datetime), [Shift]";
            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@B", beginDate.Date);
            cmd.Parameters.AddWithValue("@E", endDate.Date.AddDays(1).AddSeconds(-1));
            using var rd = cmd.ExecuteReader();

            var empHrs = new Dictionary<string, decimal>();
            var rows = new List<object>();
            while (rd.Read())
            {
                string emp  = rd["EmployeeName"].ToString() ?? "";
                DateTime sd = TryDate(rd["ShiftDate"], beginDate);
                DateTime? tIn  = BuildDt(sd, rd["TimeIn"]);
                DateTime? tOut = BuildDt(sd, rd["TimeOut"]);
                DateTime? bkO  = BuildDt(sd, rd["BreakOut"]);
                DateTime? bkI  = BuildDt(sd, rd["BreakIn"]);

                decimal hrs = CalcHours(tIn, tOut, bkO, bkI);
                if (!empHrs.ContainsKey(emp)) empHrs[emp] = 0;
                var (reg, ot) = SplitOT(empHrs[emp], hrs);
                empHrs[emp] += hrs;
                decimal rate = Convert.ToDecimal(rd["RateOfPay"]);

                rows.Add(new {
                    employeeName   = emp,
                    departmentName = rd["DepartmentName"].ToString(),
                    shiftDate      = sd.ToString("MM/dd/yyyy"),
                    totalHours     = hrs,
                    regHours       = Math.Round(reg,2),
                    otHours        = Math.Round(ot,2),
                    rateOfPay      = rate,
                    wages          = Math.Round((reg*rate)+(ot*rate*1.5m),2),
                    paidBy         = Convert.ToInt32(rd["PaidBy"])
                });
            }
            return Json(rows);
        }

        // ── 2. Overtime Watch ────────────────────────────────────────────────
        // Show employees with hours and flag those at or over OT threshold
        [HttpGet]
        public IActionResult GetOvertimeWatch(DateTime beginDate, DateTime endDate)
        {
            using var cn = Cn(); cn.Open();
            var sql = @"
                SELECT [Employee Name] AS EmployeeName,
                       [Department Name] AS DepartmentName,
                       [Shift Date] AS ShiftDate,
                       [Time In] AS TimeIn, [Time Out] AS TimeOut,
                       [Break Out] AS BreakOut, [Break In] AS BreakIn,
                       ISNULL([Rate of Pay],0) AS RateOfPay
                FROM [MSaleData].[dbo].[Employee Records]
                WHERE TRY_CAST([Shift Date] AS datetime) BETWEEN @B AND @E
                ORDER BY [Employee Name], TRY_CAST([Shift Date] AS datetime)";
            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@B", beginDate.Date);
            cmd.Parameters.AddWithValue("@E", endDate.Date.AddDays(1).AddSeconds(-1));
            using var rd = cmd.ExecuteReader();

            var empData = new Dictionary<string, (decimal hrs, decimal rate, string dept)>();
            while (rd.Read())
            {
                string emp  = rd["EmployeeName"].ToString() ?? "";
                DateTime sd = TryDate(rd["ShiftDate"], beginDate);
                DateTime? tIn  = BuildDt(sd, rd["TimeIn"]);
                DateTime? tOut = BuildDt(sd, rd["TimeOut"]);
                DateTime? bkO  = BuildDt(sd, rd["BreakOut"]);
                DateTime? bkI  = BuildDt(sd, rd["BreakIn"]);
                decimal hrs = CalcHours(tIn, tOut, bkO, bkI);
                decimal rate = Convert.ToDecimal(rd["RateOfPay"]);
                string dept = rd["DepartmentName"].ToString() ?? "";
                if (!empData.ContainsKey(emp)) empData[emp] = (0, rate, dept);
                empData[emp] = (empData[emp].hrs + hrs, rate, dept);
            }

            var result = empData.Select(kv => new {
                employeeName   = kv.Key,
                departmentName = kv.Value.dept,
                totalHours     = Math.Round(kv.Value.hrs, 2),
                otHours        = Math.Round(Math.Max(0, kv.Value.hrs - 40), 2),
                rateOfPay      = kv.Value.rate,
                otWages        = Math.Round(Math.Max(0, kv.Value.hrs - 40) * kv.Value.rate * 1.5m, 2),
                status         = kv.Value.hrs >= 40 ? "OT" : kv.Value.hrs >= 35 ? "Warning" : "OK"
            }).OrderByDescending(r => r.totalHours).ToList();

            return Json(result);
        }

        // ── 3. Hourly Labor Report ───────────────────────────────────────────
        // HourlyLaborData SDate — single day, hourly staffing + wages
        [HttpGet]
        public IActionResult GetHourlyLabor(DateTime date)
        {
            using var cn = Cn(); cn.Open();
            var sql = @"
                SELECT [Employee Name] AS EmployeeName,
                       [Department Name] AS DepartmentName,
                       [Time In] AS TimeIn, [Time Out] AS TimeOut,
                       [Break Out] AS BreakOut, [Break In] AS BreakIn,
                       ISNULL([Rate of Pay],0) AS RateOfPay,
                       ISNULL([Paid By],1) AS PaidBy
                FROM [MSaleData].[dbo].[Employee Records]
                WHERE TRY_CAST([Shift Date] AS datetime) = @D
                  AND ISNULL([Paid By],1) <> 0
                ORDER BY TRY_CAST([Time In] AS datetime)";
            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@D", date.Date);
            using var rd = cmd.ExecuteReader();

            // Build 24 hourly buckets
            var buckets = Enumerable.Range(0, 24).Select(h => new {
                hour = h,
                label = new DateTime(date.Year,date.Month,date.Day,h,0,0).ToString("hh:mm tt"),
                employees = new List<string>(),
                wages = 0m
            }).ToList();

            var shifts = new List<(string emp, string dept, DateTime? tIn, DateTime? tOut, decimal rate, decimal hrs)>();
            while (rd.Read())
            {
                string emp  = rd["EmployeeName"].ToString() ?? "";
                DateTime? tIn  = BuildDt(date, rd["TimeIn"]);
                DateTime? tOut = BuildDt(date, rd["TimeOut"]);
                DateTime? bkO  = BuildDt(date, rd["BreakOut"]);
                DateTime? bkI  = BuildDt(date, rd["BreakIn"]);
                decimal rate = Convert.ToDecimal(rd["RateOfPay"]);
                decimal hrs = CalcHours(tIn, tOut, bkO, bkI);
                shifts.Add((emp, rd["DepartmentName"].ToString() ?? "", tIn, tOut, rate, hrs));
            }

            var hourlyRows = Enumerable.Range(0, 24).Select(h => {
                var hStart = new DateTime(date.Year,date.Month,date.Day,h,0,0);
                var hEnd   = hStart.AddHours(1);
                var onClock = shifts.Where(s =>
                    s.tIn.HasValue && s.tOut.HasValue &&
                    s.tIn.Value < hEnd && s.tOut.Value > hStart).ToList();
                decimal wages = onClock.Sum(s => s.rate);
                return new {
                    hour       = h,
                    label      = hStart.ToString("hh:mm tt"),
                    headCount  = onClock.Count,
                    employees  = onClock.Select(s => s.emp).ToList(),
                    laborCost  = Math.Round(wages, 2)
                };
            }).ToList();

            return Json(new { date = date.ToString("MM/dd/yyyy"), hours = hourlyRows, shifts });
        }

        // ── 4. Punch Report ──────────────────────────────────────────────────
        // Per-employee raw punch in/out records
        [HttpGet]
        public IActionResult GetPunchReport(DateTime beginDate, DateTime endDate)
        {
            using var cn = Cn(); cn.Open();
            var sql = @"
                SELECT [Employee Name] AS EmployeeName,
                       [Department Name] AS DepartmentName,
                       [Shift Date] AS ShiftDate,
                       [Punch Out Date] AS PunchOutDate,
                       [Time In] AS TimeIn, [Time Out] AS TimeOut,
                       [Break Out] AS BreakOut, [Break In] AS BreakIn,
                       ISNULL([Altered By],'') AS AlteredBy,
                       ISNULL([Altered In],0)  AS AlteredIn,
                       ISNULL([Altered Out],0) AS AlteredOut,
                       ISNULL([Shift],0) AS Shift
                FROM [MSaleData].[dbo].[Employee Records]
                WHERE TRY_CAST([Shift Date] AS datetime) BETWEEN @B AND @E
                ORDER BY [Employee Name], TRY_CAST([Shift Date] AS datetime),
                         TRY_CAST([Time In] AS datetime)";
            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@B", beginDate.Date);
            cmd.Parameters.AddWithValue("@E", endDate.Date.AddDays(1).AddSeconds(-1));
            using var rd = cmd.ExecuteReader();

            var rows = new List<object>();
            while (rd.Read())
            {
                DateTime sd = TryDate(rd["ShiftDate"], beginDate);
                DateTime pd = TryDate(rd["PunchOutDate"], sd);
                rows.Add(new {
                    employeeName   = rd["EmployeeName"].ToString(),
                    departmentName = rd["DepartmentName"].ToString(),
                    shiftDate      = sd.ToString("MM/dd/yyyy"),
                    punchOutDate   = pd.ToString("MM/dd/yyyy"),
                    timeIn         = FmtTime(rd["TimeIn"]),
                    timeOut        = FmtTime(rd["TimeOut"]),
                    breakOut       = FmtTime(rd["BreakOut"]),
                    breakIn        = FmtTime(rd["BreakIn"]),
                    alteredBy      = rd["AlteredBy"].ToString(),
                    alteredIn      = Convert.ToInt32(rd["AlteredIn"]),
                    alteredOut     = Convert.ToInt32(rd["AlteredOut"])
                });
            }
            return Json(rows);
        }

        // ── 5. Who's On the Clock ────────────────────────────────────────────
        // Employees currently punched in (no Time Out)
        [HttpGet]
        public IActionResult GetWhoIsOnClock()
        {
            using var cn = Cn(); cn.Open();
            var sql = @"
                SELECT DISTINCT [Employee Name] AS EmployeeName,
                       [Department Name] AS DepartmentName,
                       [Time In] AS TimeIn,
                       [Shift Date] AS ShiftDate
                FROM [MSaleData].[dbo].[Employee Records]
                WHERE [Time In] > ''
                  AND (ISDATE([Time Out]) = 0 OR [Time Out] IS NULL OR LTRIM(RTRIM([Time Out])) = '')
                ORDER BY [Employee Name]";
            using var cmd = new SqlCommand(sql, cn);
            using var rd = cmd.ExecuteReader();
            var rows = new List<object>();
            while (rd.Read())
            {
                DateTime sd = TryDate(rd["ShiftDate"], DateTime.Today);
                TimeSpan elapsed = DateTime.Now - (BuildDt(sd, rd["TimeIn"]) ?? DateTime.Now);
                rows.Add(new {
                    employeeName   = rd["EmployeeName"].ToString(),
                    departmentName = rd["DepartmentName"].ToString(),
                    shiftDate      = sd.ToString("MM/dd/yyyy"),
                    timeIn         = FmtTime(rd["TimeIn"]),
                    hoursOn        = Math.Round(elapsed.TotalHours, 2)
                });
            }
            return Json(rows);
        }

        // ── 6. Break Report ──────────────────────────────────────────────────
        [HttpGet]
        public IActionResult GetBreakReport(DateTime beginDate, DateTime endDate)
        {
            using var cn = Cn(); cn.Open();
            var sql = @"
                SELECT [Employee Name] AS EmployeeName,
                       [Department Name] AS DepartmentName,
                       [Shift Date] AS ShiftDate,
                       [Time In] AS TimeIn, [Time Out] AS TimeOut,
                       [Break Out] AS BreakOut, [Break In] AS BreakIn
                FROM [MSaleData].[dbo].[Employee Records]
                WHERE TRY_CAST([Shift Date] AS date) BETWEEN @B AND @E
                  AND [Break Out] IS NOT NULL AND ISDATE([Break Out]) = 1
                ORDER BY TRY_CAST([Shift Date] AS date), [Employee Name]";
            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@B", beginDate.Date);
            cmd.Parameters.AddWithValue("@E", endDate.Date);
            using var rd = cmd.ExecuteReader();

            var rows = new List<object>();
            while (rd.Read())
            {
                DateTime sd  = TryDate(rd["ShiftDate"], beginDate);
                DateTime? bkO = BuildDt(sd, rd["BreakOut"]);
                DateTime? bkI = BuildDt(sd, rd["BreakIn"]);
                decimal mins  = 0;
                if (bkO.HasValue && bkI.HasValue && bkI > bkO)
                    mins = Math.Round((decimal)(bkI.Value - bkO.Value).TotalMinutes, 1);
                rows.Add(new {
                    employeeName   = rd["EmployeeName"].ToString(),
                    departmentName = rd["DepartmentName"].ToString(),
                    shiftDate      = sd.ToString("MM/dd/yyyy"),
                    timeIn         = FmtTime(rd["TimeIn"]),
                    timeOut        = FmtTime(rd["TimeOut"]),
                    breakOut       = FmtTime(rd["BreakOut"]),
                    breakIn        = FmtTime(rd["BreakIn"]),
                    breakMinutes   = mins
                });
            }
            return Json(rows);
        }

        // ── 7. Hours By Department ───────────────────────────────────────────
        // DepartmentIndReport BeginWeek, EndingDate
        [HttpGet]
        public IActionResult GetHoursByDepartment(DateTime beginDate, DateTime endDate)
        {
            using var cn = Cn(); cn.Open();
            var sql = @"
                SELECT [Employee Name] AS EmployeeName,
                       [Department Name] AS DepartmentName,
                       [Shift Date] AS ShiftDate,
                       [Time In] AS TimeIn, [Time Out] AS TimeOut,
                       [Break Out] AS BreakOut, [Break In] AS BreakIn,
                       ISNULL([Rate of Pay],0) AS RateOfPay
                FROM [MSaleData].[dbo].[Employee Records]
                WHERE TRY_CAST([Shift Date] AS datetime) BETWEEN @B AND @E
                ORDER BY [Department Name], TRY_CAST([Shift Date] AS datetime), [Employee Name]";
            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@B", beginDate.Date);
            cmd.Parameters.AddWithValue("@E", endDate.Date.AddDays(1).AddSeconds(-1));
            using var rd = cmd.ExecuteReader();

            // Group by dept > employee
            var depts = new Dictionary<string, Dictionary<string, (decimal hrs, decimal wages, int shifts)>>();
            while (rd.Read())
            {
                string dept = rd["DepartmentName"].ToString() ?? "(None)";
                string emp  = rd["EmployeeName"].ToString() ?? "";
                DateTime sd = TryDate(rd["ShiftDate"], beginDate);
                DateTime? tIn  = BuildDt(sd, rd["TimeIn"]);
                DateTime? tOut = BuildDt(sd, rd["TimeOut"]);
                DateTime? bkO  = BuildDt(sd, rd["BreakOut"]);
                DateTime? bkI  = BuildDt(sd, rd["BreakIn"]);
                decimal hrs  = CalcHours(tIn, tOut, bkO, bkI);
                decimal rate = Convert.ToDecimal(rd["RateOfPay"]);
                if (!depts.ContainsKey(dept)) depts[dept] = new();
                if (!depts[dept].ContainsKey(emp)) depts[dept][emp] = (0,0,0);
                var cur = depts[dept][emp];
                depts[dept][emp] = (cur.hrs+hrs, cur.wages+(hrs*rate), cur.shifts+1);
            }

            var result = depts.OrderBy(d=>d.Key).Select(d => new {
                departmentName = d.Key,
                totalHours = Math.Round(d.Value.Sum(e=>e.Value.hrs),2),
                totalWages = Math.Round(d.Value.Sum(e=>e.Value.wages),2),
                employees  = d.Value.OrderBy(e=>e.Key).Select(e => new {
                    employeeName = e.Key,
                    totalHours   = Math.Round(e.Value.hrs,2),
                    totalWages   = Math.Round(e.Value.wages,2),
                    shifts       = e.Value.shifts
                }).ToList()
            }).ToList();
            return Json(result);
        }

        // ── 8. Avg Hourly Rate Exception (Tipped Employees) ──────────────────
        // AverageRateofPay — employees where tip avg + rate is below min wage
        [HttpGet]
        public IActionResult GetAvgRateException(DateTime beginDate, DateTime endDate)
        {
            using var cn = Cn(); cn.Open();
            var sql = @"
                SELECT [Employee Name] AS EmployeeName,
                       [Department Name] AS DepartmentName,
                       [Shift Date] AS ShiftDate,
                       [Time In] AS TimeIn, [Time Out] AS TimeOut,
                       [Break Out] AS BreakOut, [Break In] AS BreakIn,
                       ISNULL([Rate of Pay],0) AS RateOfPay,
                       ISNULL([Tip Amount],0) AS DeclaredTips,
                       ISNULL([Credit Card Tips],0) AS CreditCardTips,
                       ISNULL([Sales],0) AS Sales
                FROM [MSaleData].[dbo].[Employee Records]
                WHERE TRY_CAST([Shift Date] AS datetime) BETWEEN @B AND @E
                  AND (ISNULL([Tip Amount],0) > 0 OR ISNULL([Credit Card Tips],0) > 0)
                ORDER BY [Employee Name], TRY_CAST([Shift Date] AS datetime)";
            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@B", beginDate.Date);
            cmd.Parameters.AddWithValue("@E", endDate.Date.AddDays(1).AddSeconds(-1));
            using var rd = cmd.ExecuteReader();

            var empData = new Dictionary<string, (decimal hrs, decimal rate, decimal tips, decimal sales, int shifts)>();
            while (rd.Read())
            {
                string emp  = rd["EmployeeName"].ToString() ?? "";
                DateTime sd = TryDate(rd["ShiftDate"], beginDate);
                DateTime? tIn  = BuildDt(sd, rd["TimeIn"]);
                DateTime? tOut = BuildDt(sd, rd["TimeOut"]);
                DateTime? bkO  = BuildDt(sd, rd["BreakOut"]);
                DateTime? bkI  = BuildDt(sd, rd["BreakIn"]);
                decimal hrs  = CalcHours(tIn, tOut, bkO, bkI);
                decimal rate = Convert.ToDecimal(rd["RateOfPay"]);
                decimal tips = Convert.ToDecimal(rd["DeclaredTips"]) + Convert.ToDecimal(rd["CreditCardTips"]);
                decimal sales = Convert.ToDecimal(rd["Sales"]);
                if (!empData.ContainsKey(emp)) empData[emp] = (0,rate,0,0,0);
                var cur = empData[emp];
                empData[emp] = (cur.hrs+hrs, rate, cur.tips+tips, cur.sales+sales, cur.shifts+1);
            }

            decimal minWage = 7.25m; // federal minimum — adjust via config if needed
            var result = empData.Select(kv => {
                decimal avgRate = kv.Value.hrs > 0
                    ? Math.Round((kv.Value.rate * kv.Value.hrs + kv.Value.tips) / kv.Value.hrs, 2)
                    : kv.Value.rate;
                return new {
                    employeeName = kv.Key,
                    totalHours   = Math.Round(kv.Value.hrs, 2),
                    baseRate     = kv.Value.rate,
                    totalTips    = Math.Round(kv.Value.tips, 2),
                    totalSales   = Math.Round(kv.Value.sales, 2),
                    avgHourlyRate = avgRate,
                    belowMinWage = avgRate < minWage,
                    shifts       = kv.Value.shifts
                };
            }).OrderBy(r => r.avgHourlyRate).ToList();
            return Json(result);
        }

        // ── 9. Individual Daily Time & Sales ─────────────────────────────────
        // GetTimeReport BeginWeek, EndingDate — per employee, per day
        [HttpGet]
        public IActionResult GetIndividualDailyReport(DateTime beginDate, DateTime endDate, string? employeeName)
        {
            using var cn = Cn(); cn.Open();
            var sql = @"
                SELECT [Employee Name] AS EmployeeName,
                       [Department Name] AS DepartmentName,
                       [Shift Date] AS ShiftDate,
                       [Time In] AS TimeIn, [Time Out] AS TimeOut,
                       [Break Out] AS BreakOut, [Break In] AS BreakIn,
                       ISNULL([Rate of Pay],0) AS RateOfPay,
                       ISNULL([Tip Amount],0) AS DeclaredTips,
                       ISNULL([Credit Card Tips],0) AS CreditCardTips,
                       ISNULL([Tips Paid Out],0) AS TipsPaidOut,
                       ISNULL([Tip Share],0) AS TipShare,
                       ISNULL([Sales],0) AS Sales,
                       ISNULL([Altered By],'') AS AlteredBy
                FROM [MSaleData].[dbo].[Employee Records]
                WHERE TRY_CAST([Shift Date] AS datetime) BETWEEN @B AND @E
                  AND (@Emp IS NULL OR [Employee Name] = @Emp)
                  AND ISDATE([Time In]) = 1
                ORDER BY [Employee Name], TRY_CAST([Shift Date] AS datetime),
                         TRY_CAST([Time In] AS datetime)";
            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@B",   beginDate.Date);
            cmd.Parameters.AddWithValue("@E",   endDate.Date.AddDays(1).AddSeconds(-1));
            cmd.Parameters.AddWithValue("@Emp", string.IsNullOrWhiteSpace(employeeName) ? (object)DBNull.Value : employeeName);
            using var rd = cmd.ExecuteReader();

            var empHrs = new Dictionary<string, decimal>();
            var rows   = new List<object>();
            while (rd.Read())
            {
                string emp  = rd["EmployeeName"].ToString() ?? "";
                DateTime sd = TryDate(rd["ShiftDate"], beginDate);
                DateTime? tIn  = BuildDt(sd, rd["TimeIn"]);
                DateTime? tOut = BuildDt(sd, rd["TimeOut"]);
                DateTime? bkO  = BuildDt(sd, rd["BreakOut"]);
                DateTime? bkI  = BuildDt(sd, rd["BreakIn"]);
                decimal hrs = CalcHours(tIn, tOut, bkO, bkI);
                if (!empHrs.ContainsKey(emp)) empHrs[emp] = 0;
                var (reg, ot) = SplitOT(empHrs[emp], hrs);
                empHrs[emp] += hrs;
                decimal rate = Convert.ToDecimal(rd["RateOfPay"]);
                rows.Add(new {
                    employeeName   = emp,
                    departmentName = rd["DepartmentName"].ToString(),
                    shiftDate      = sd.ToString("MM/dd/yyyy"),
                    timeIn         = FmtTime(rd["TimeIn"]),
                    timeOut        = FmtTime(rd["TimeOut"]),
                    totalHours     = hrs,
                    regHours       = Math.Round(reg,2),
                    otHours        = Math.Round(ot,2),
                    rateOfPay      = rate,
                    wages          = Math.Round((reg*rate)+(ot*rate*1.5m),2),
                    declaredTips   = Convert.ToDecimal(rd["DeclaredTips"]),
                    creditCardTips = Convert.ToDecimal(rd["CreditCardTips"]),
                    tipsPaidOut    = Convert.ToDecimal(rd["TipsPaidOut"]),
                    tipShare       = Convert.ToDecimal(rd["TipShare"]),
                    sales          = Convert.ToDecimal(rd["Sales"]),
                    alteredBy      = rd["AlteredBy"].ToString()
                });
            }
            return Json(rows);
        }

        // ── 10. Employee List (for dropdowns) ────────────────────────────────
        [HttpGet]
        public IActionResult GetEmployeeList(DateTime beginDate, DateTime endDate)
        {
            using var cn = Cn();
            var names = cn.Query<string>(@"
                SELECT DISTINCT [Employee Name] FROM [MSaleData].[dbo].[Employee Records]
                WHERE TRY_CAST([Shift Date] AS datetime) BETWEEN @B AND @E
                ORDER BY [Employee Name]",
                new { B = beginDate.Date, E = endDate.Date.AddDays(1).AddSeconds(-1) });
            return Json(names);
        }

        // ── 11. Phone List ───────────────────────────────────────────────────
        [HttpGet]
        public IActionResult GetPhoneList()
        {
            using var cn = Cn(); cn.Open();
            var sql = @"
                SELECT [First Name] AS FirstName, [Last Name] AS LastName,
                       ISNULL([Home Phone],'') AS HomePhone,
                       ISNULL([Phone # 2],'')  AS Phone2
                FROM [MSaleData].[dbo].[General Data]
                WHERE ISNULL([Terminated],'0') = '0' OR [Terminated] = ''
                ORDER BY [Last Name], [First Name]";
            try {
                using var cmd = new SqlCommand(sql, cn);
                using var rd = cmd.ExecuteReader();
                var rows = new List<object>();
                while (rd.Read())
                    rows.Add(new {
                        firstName = rd["FirstName"].ToString(),
                        lastName  = rd["LastName"].ToString(),
                        homePhone = rd["HomePhone"].ToString(),
                        phone2    = rd["Phone2"].ToString()
                    });
                return Json(rows);
            } catch {
                return Json(new List<object>());
            }
        }

        // ── 12. Access Codes ─────────────────────────────────────────────────
        [HttpGet]
        public IActionResult GetAccessCodes()
        {
            using var cn = Cn(); cn.Open();
            var sql = @"
                SELECT [First Name] AS FirstName, [Last Name] AS LastName,
                       ISNULL([Access Level],'')  AS AccessLevel,
                       ISNULL([Access Number],'') AS AccessNumber
                FROM [MSaleData].[dbo].[General Data]
                WHERE ISNULL([Terminated],'0') = '0' OR [Terminated] = ''
                ORDER BY [Last Name], [First Name]";
            try {
                using var cmd = new SqlCommand(sql, cn);
                using var rd = cmd.ExecuteReader();
                var rows = new List<object>();
                while (rd.Read())
                    rows.Add(new {
                        firstName    = rd["FirstName"].ToString(),
                        lastName     = rd["LastName"].ToString(),
                        accessLevel  = rd["AccessLevel"].ToString(),
                        accessNumber = rd["AccessNumber"].ToString()
                    });
                return Json(rows);
            } catch {
                return Json(new List<object>());
            }
        }

        // ── Helpers ──────────────────────────────────────────────────────────
        static DateTime TryDate(object v, DateTime fallback) =>
            v != DBNull.Value && DateTime.TryParse(v?.ToString(), out var d) ? d : fallback;

        static DateTime? BuildDt(DateTime base_, object v) {
            if (v == DBNull.Value || v == null) return null;
            if (!DateTime.TryParse(v.ToString(), out var t)) return null;
            return base_.Date.Add(t.TimeOfDay);
        }

        static string FmtTime(object v) {
            if (v == DBNull.Value || v == null) return "";
            if (!DateTime.TryParse(v.ToString(), out var t)) return "";
            return t.ToString("hh:mm tt");
        }

        static decimal CalcHours(DateTime? tIn, DateTime? tOut, DateTime? bkO, DateTime? bkI) {
            if (!tIn.HasValue || !tOut.HasValue || tOut <= tIn) return 0;
            decimal hrs = (decimal)(tOut.Value - tIn.Value).TotalHours;
            if (bkO.HasValue && bkI.HasValue && bkI > bkO)
                hrs -= (decimal)(bkI.Value - bkO.Value).TotalHours;
            return Math.Max(0, Math.Round(hrs, 2));
        }

        static (decimal reg, decimal ot) SplitOT(decimal before, decimal hrs) {
            decimal after = before + hrs;
            if (before >= 40) return (0, hrs);
            if (after > 40)   return (40 - before, after - 40);
            return (hrs, 0);
        }
    }
}
