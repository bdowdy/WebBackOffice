using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MSaleWebServer.Models;
using System.Configuration;

namespace MSaleWebServer.Controllers
{
    public class TimeRecordsEditor : Controller
    {
        private readonly IConfiguration _configuration;
        public TimeRecordsEditor(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public IActionResult GetTimeRecords(DateTime beginDate, DateTime endDate, string? employeeName)
        {
            var rows = new List<TimeRecordDto>();
            var cs = _configuration.GetConnectionString("DefaultConnection");

            using var cn = new SqlConnection(cs);
            cn.Open();

            string sql = @"
        SELECT
            [RecordId],
            [Employee Name],
            [Department Name],
            [Shift Date],
            [Time In],
            [Break Out],
            [Break In],
            [Time Out],
            [Punch Out Date],
            ISNULL([Tip Amount],0) AS TipAmount,
            ISNULL([Rate of Pay],0) AS RateOfPay,
            ISNULL([Sales],0) AS Sales,
            ISNULL([Credit Card Tips],0) AS CreditCardTips,
            ISNULL([Tips Paid Out],0) AS TipsPaidOut,
            ISNULL([Tip Share],0) AS TipShare,
            ISNULL([Shift],0) AS Shift
        FROM [MSaleData].[dbo].[Employee Records]
        WHERE TRY_CAST([Shift Date] AS datetime) BETWEEN @BeginDate AND @EndDate
          AND (@EmployeeName = '' OR [Employee Name] = @EmployeeName)
        ORDER BY [Employee Name], TRY_CAST([Shift Date] AS datetime), TRY_CAST([Time In] AS datetime)";

            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@BeginDate", beginDate.Date);
            cmd.Parameters.AddWithValue("@EndDate", endDate.Date.AddDays(1).AddSeconds(-1));
            cmd.Parameters.AddWithValue("@EmployeeName", employeeName ?? "");

            using var rd = cmd.ExecuteReader();

            while (rd.Read())
            {
                rows.Add(new TimeRecordDto
                {
                    RecordId = rd["RecordId"].ToString() ?? "",
                    EmployeeName = rd["Employee Name"].ToString() ?? "",
                    DepartmentName = rd["Department Name"].ToString() ?? "",
                    ShiftDate = FormatDate(rd["Shift Date"]),
                    TimeIn = FormatTime(rd["Time In"]),
                    BreakOut = FormatTime(rd["Break Out"]),
                    BreakIn = FormatTime(rd["Break In"]),
                    TimeOut = FormatTime(rd["Time Out"]),
                    PunchOutDate = FormatDate(rd["Punch Out Date"]),
                    TipAmount = Convert.ToDecimal(rd["TipAmount"]),
                    RateOfPay = Convert.ToDecimal(rd["RateOfPay"]),
                    Sales = Convert.ToDecimal(rd["Sales"]),
                    CreditCardTips = Convert.ToDecimal(rd["CreditCardTips"]),
                    TipsPaidOut = Convert.ToDecimal(rd["TipsPaidOut"]),
                    TipShare = Convert.ToDecimal(rd["TipShare"]),
                    Shift = Convert.ToInt32(rd["Shift"])
                });
            }

            return Json(rows);
        }

        [HttpGet]
        public IActionResult GetTimeRecordEmployees()
        {
            var list = new List<string>();
            var cs = _configuration.GetConnectionString("DefaultConnection");

            using var cn = new SqlConnection(cs);
            cn.Open();

            string sql = @"
        SELECT DISTINCT [Employee Name]
        FROM [MSaleData].[dbo].[Employee Records]
        WHERE ISNULL([Employee Name],'') <> ''
        ORDER BY [Employee Name]";

            using var cmd = new SqlCommand(sql, cn);
            using var rd = cmd.ExecuteReader();

            while (rd.Read())
                list.Add(rd[0].ToString() ?? "");

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveTimeRecord([FromBody] TimeRecordDto item)
        {
            var cs = _configuration.GetConnectionString("DefaultConnection");

            if (string.IsNullOrWhiteSpace(item.RecordId))
                item.RecordId = Guid.NewGuid().ToString();

            using var cn = new SqlConnection(cs);
            cn.Open();

            string existsSql = @"
        SELECT COUNT(*)
        FROM [MSaleData].[dbo].[Employee Records]
        WHERE [RecordId] = @RecordId";

            using var existsCmd = new SqlCommand(existsSql, cn);
            existsCmd.Parameters.AddWithValue("@RecordId", item.RecordId);

            bool exists = Convert.ToInt32(existsCmd.ExecuteScalar()) > 0;

            string sql = exists ? @"
        UPDATE [MSaleData].[dbo].[Employee Records]
        SET
            [Employee Name] = @EmployeeName,
            [Department Name] = @DepartmentName,
            [Shift Date] = @ShiftDate,
            [Time In] = @TimeIn,
            [Break Out] = @BreakOut,
            [Break In] = @BreakIn,
            [Time Out] = @TimeOut,
            [Punch Out Date] = @PunchOutDate,
            [Tip Amount] = @TipAmount,
            [Rate of Pay] = @RateOfPay,
            [Sales] = @Sales,
            [Credit Card Tips] = @CreditCardTips,
            [Tips Paid Out] = @TipsPaidOut,
            [Tip Share] = @TipShare,
            [Shift] = @Shift,
            [Altered By] = 'MSaleWebServer'
        WHERE [RecordId] = @RecordId"
            : @"
        INSERT INTO [MSaleData].[dbo].[Employee Records]
        (
            [Employee Name], [Department Name], [Shift Date],
            [Time In], [Break Out], [Break In], [Time Out], [Punch Out Date],
            [Tip Amount], [Rate of Pay], [Sales], [Credit Card Tips],
            [Tips Paid Out], [Tip Share], [Shift], [RecordId], [Altered By]
        )
        VALUES
        (
            @EmployeeName, @DepartmentName, @ShiftDate,
            @TimeIn, @BreakOut, @BreakIn, @TimeOut, @PunchOutDate,
            @TipAmount, @RateOfPay, @Sales, @CreditCardTips,
            @TipsPaidOut, @TipShare, @Shift, @RecordId, 'MSaleWebServer'
        )";

            using var cmd = new SqlCommand(sql, cn);
            AddTimeRecordParams(cmd, item);
            cmd.ExecuteNonQuery();

            return Json(new { success = true, message = "Time record saved." });
        }

        [HttpPost]
        public IActionResult DeleteTimeRecord([FromBody] TimeRecordDto item)
        {
            var cs = _configuration.GetConnectionString("DefaultConnection");

            using var cn = new SqlConnection(cs);
            cn.Open();

            string sql = @"
        DELETE FROM [MSaleData].[dbo].[Employee Records]
        WHERE [RecordId] = @RecordId";

            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@RecordId", item.RecordId);
            cmd.ExecuteNonQuery();

            return Json(new { success = true, message = "Time record deleted." });
        }

        private static void AddTimeRecordParams(SqlCommand cmd, TimeRecordDto item)
        {
            cmd.Parameters.AddWithValue("@RecordId", item.RecordId ?? "");
            cmd.Parameters.AddWithValue("@EmployeeName", item.EmployeeName ?? "");
            cmd.Parameters.AddWithValue("@DepartmentName", item.DepartmentName ?? "");
            cmd.Parameters.AddWithValue("@ShiftDate", ToDbDate(item.ShiftDate));
            cmd.Parameters.AddWithValue("@TimeIn", ToDbTime(item.ShiftDate, item.TimeIn));
            cmd.Parameters.AddWithValue("@BreakOut", ToDbTime(item.ShiftDate, item.BreakOut));
            cmd.Parameters.AddWithValue("@BreakIn", ToDbTime(item.ShiftDate, item.BreakIn));
            cmd.Parameters.AddWithValue("@TimeOut", ToDbTime(item.PunchOutDate, item.TimeOut));
            cmd.Parameters.AddWithValue("@PunchOutDate", ToDbDate(item.PunchOutDate));
            cmd.Parameters.AddWithValue("@TipAmount", item.TipAmount);
            cmd.Parameters.AddWithValue("@RateOfPay", item.RateOfPay);
            cmd.Parameters.AddWithValue("@Sales", item.Sales);
            cmd.Parameters.AddWithValue("@CreditCardTips", item.CreditCardTips);
            cmd.Parameters.AddWithValue("@TipsPaidOut", item.TipsPaidOut);
            cmd.Parameters.AddWithValue("@TipShare", item.TipShare);
            cmd.Parameters.AddWithValue("@Shift", item.Shift);
        }

        private static object ToDbDate(string value)
        {
            if (DateTime.TryParse(value, out var d))
                return d.Date;

            return DBNull.Value;
        }

        private static object ToDbTime(string dateValue, string timeValue)
        {
            if (!DateTime.TryParse(dateValue, out var d))
                return DBNull.Value;

            if (!DateTime.TryParse(timeValue, out var t))
                return DBNull.Value;

            return d.Date.Add(t.TimeOfDay);
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
        

    }
}
