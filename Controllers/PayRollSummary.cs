using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MSaleWebServer.Models;
using System.Configuration;

namespace MSaleWebServer.Controllers
{
    public class PayRollSummary : Controller
    {
        private readonly IConfiguration _configuration;

        public PayRollSummary(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public IActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public IActionResult GetPayrollSummary(DateTime beginDate, DateTime endDate)
        {
            var rows = new List<PayrollSummaryRow>();
            var cs = _configuration.GetConnectionString("DefaultConnection");

            using var cn = new SqlConnection(cs);
            cn.Open();

            string sql = @"
        SELECT
            [Employee Name],
            [Department Name],
            [Shift Date],
            [Time In],
            [Break Out],
            [Break In],
            [Time Out],
            [Punch Out Date],
            ISNULL([Rate of Pay], 0) AS RateOfPay,
            ISNULL([Tip Amount], 0) AS DeclaredTips,
            ISNULL([Credit Card Tips], 0) AS CreditCardTips,
            ISNULL([Tips Paid Out], 0) AS TipsPaidOut,
            ISNULL([Tip Share], 0) AS TipShare,
            ISNULL([Sales], 0) AS Sales,
            ISNULL([Shift], 0) AS Shift
        FROM [MSaleData].[dbo].[Employee Records]
        WHERE TRY_CAST([Shift Date] AS datetime) BETWEEN @BeginDate AND @EndDate
        ORDER BY [Employee Name], TRY_CAST([Shift Date] AS datetime), [Shift]";

            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@BeginDate", beginDate.Date);
            cmd.Parameters.AddWithValue("@EndDate", endDate.Date.AddDays(1).AddSeconds(-1));

            using var rd = cmd.ExecuteReader();

            var employeeHours = new Dictionary<string, decimal>();

            while (rd.Read())
            {
                string empName = rd["Employee Name"].ToString() ?? "";
                string deptName = rd["Department Name"].ToString() ?? "";

                DateTime shiftDate = Convert.ToDateTime(rd["Shift Date"]);
                DateTime punchOutDate = rd["Punch Out Date"] == DBNull.Value
                    ? shiftDate
                    : Convert.ToDateTime(rd["Punch Out Date"]);

                DateTime? timeIn = BuildDateTime(shiftDate, rd["Time In"]);
                DateTime? timeOut = BuildDateTime(punchOutDate, rd["Time Out"]);

                decimal totalHours = 0;

                if (timeIn.HasValue && timeOut.HasValue && timeOut.Value > timeIn.Value)
                {
                    totalHours = (decimal)(timeOut.Value - timeIn.Value).TotalHours;
                }

                DateTime? breakOut = BuildDateTime(shiftDate, rd["Break Out"]);
                DateTime? breakIn = BuildDateTime(shiftDate, rd["Break In"]);

                if (breakOut.HasValue && breakIn.HasValue && breakIn.Value > breakOut.Value)
                {
                    totalHours -= (decimal)(breakIn.Value - breakOut.Value).TotalHours;
                }

                if (totalHours < 0)
                    totalHours = 0;

                totalHours = Math.Round(totalHours, 2);

                if (!employeeHours.ContainsKey(empName))
                    employeeHours[empName] = 0;

                decimal beforeHours = employeeHours[empName];
                decimal afterHours = beforeHours + totalHours;

                decimal regHours = 0;
                decimal otHours = 0;

                if (beforeHours >= 40)
                {
                    otHours = totalHours;
                }
                else if (afterHours > 40)
                {
                    regHours = 40 - beforeHours;
                    otHours = afterHours - 40;
                }
                else
                {
                    regHours = totalHours;
                }

                employeeHours[empName] = afterHours;

                decimal rate = Convert.ToDecimal(rd["RateOfPay"]);
                decimal wages = (regHours * rate) + (otHours * rate * 1.5m);

                rows.Add(new PayrollSummaryRow
                {
                    EmployeeName = empName,
                    DepartmentName = deptName,
                    ShiftDate = shiftDate,
                    RateOfPay = rate,
                    RegHours = regHours,
                    OTHours = otHours,
                    TotalHours = totalHours,
                    Wages = wages,
                    DeclaredTips = Convert.ToDecimal(rd["DeclaredTips"]),
                    CreditCardTips = Convert.ToDecimal(rd["CreditCardTips"]),
                    TipsPaidOut = Convert.ToDecimal(rd["TipsPaidOut"]),
                    TipShare = Convert.ToDecimal(rd["TipShare"]),
                    Sales = Convert.ToDecimal(rd["Sales"])
                });
            }

            return Json(rows);
        }

        private static DateTime? BuildDateTime(DateTime baseDate, object value)
        {
            if (value == DBNull.Value || value == null)
                return null;

            if (!DateTime.TryParse(value.ToString(), out DateTime parsed))
                return null;

            return baseDate.Date.Add(parsed.TimeOfDay);
        }
    }
}
