using Microsoft.AspNetCore.Mvc;
using MSaleWebServer.Models;
using System.Configuration;

namespace MSaleWebServer.Controllers
{
    public class  TipEditor : Controller
    {
        private readonly IConfiguration _configuration;
        public TipEditor(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public IActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public IActionResult GetTipEmployees()
        {
            try
            {
                var list = new List<string>();
                var cs = _configuration.GetConnectionString("DefaultConnection");

                using var cn = new Microsoft.Data.SqlClient.SqlConnection(cs);
                cn.Open();

                string sql = @"
            SELECT DISTINCT [Employee Name]
            FROM [dbo].[Employee Records]
            WHERE ISNULL([Employee Name], '') <> ''
              AND TRY_CAST([Shift Date] AS datetime) >= DATEADD(day, -25, GETDATE())
            ORDER BY [Employee Name]";

                using var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, cn);
                using var rd = cmd.ExecuteReader();

                while (rd.Read())
                    list.Add(rd[0].ToString() ?? "");

                return Json(new { success = true, data = list });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public IActionResult GetTipEditRecords(string employeeName)
        {
            try
            {
                var rows = new List<TipEditRow>();
                var cs = _configuration.GetConnectionString("DefaultConnection");

                using var cn = new Microsoft.Data.SqlClient.SqlConnection(cs);
                cn.Open();

                string sql = @"
            SELECT
                er.[RecordId],
                er.[Employee Name],
                er.[Department Name],
                er.[Shift Date],
                er.[Time In],
                er.[Time Out],
                ISNULL(er.[Tip Amount], 0) AS CashTips,
                ISNULL(er.[Credit Card Tips], 0) AS CreditCardTips,
                ISNULL(er.[Tips Paid Out], 0) AS TipsPaidOut,
                ISNULL(er.[Tip Share], 0) AS TipShare,
                ISNULL(er.[Shift], 0) AS Shift,
                ISNULL(ti.[Gratuity], 0) AS Gratuity
            FROM [dbo].[Employee Records] er
            LEFT JOIN [dbo].[Tip Income] ti
                ON ti.[Server Name] = er.[Employee Name]
               AND TRY_CAST(ti.[Date] AS date) = TRY_CAST(er.[Shift Date] AS date)
               AND ISNULL(ti.[Shift], 0) = ISNULL(er.[Shift], 0)
            WHERE er.[Employee Name] = @EmployeeName
              AND TRY_CAST(er.[Shift Date] AS datetime) >= DATEADD(day, -25, GETDATE())
            ORDER BY TRY_CAST(er.[Shift Date] AS datetime), TRY_CAST(er.[Time In] AS datetime)";

                using var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, cn);
                cmd.Parameters.AddWithValue("@EmployeeName", employeeName ?? "");

                using var rd = cmd.ExecuteReader();

                while (rd.Read())
                {
                    rows.Add(new TipEditRow
                    {
                        RecordId = rd["RecordId"].ToString() ?? "",
                        EmployeeName = rd["Employee Name"].ToString() ?? "",
                        DepartmentName = rd["Department Name"].ToString() ?? "",
                        ShiftDate = FormatDate(rd["Shift Date"]),
                        TimeIn = FormatTime(rd["Time In"]),
                        TimeOut = FormatTime(rd["Time Out"]),
                        CashTips = Convert.ToDecimal(rd["CashTips"]),
                        CreditCardTips = Convert.ToDecimal(rd["CreditCardTips"]),
                        TipsPaidOut = Convert.ToDecimal(rd["TipsPaidOut"]),
                        TipShare = Convert.ToDecimal(rd["TipShare"]),
                        Gratuity = Convert.ToDecimal(rd["Gratuity"]),
                        Shift = Convert.ToInt32(rd["Shift"])
                    });
                }

                return Json(new { success = true, data = rows });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public IActionResult SaveTipEditRecords([FromBody] List<TipEditRow> rows)
        {
            try
            {
                var cs = _configuration.GetConnectionString("DefaultConnection");

                using var cn = new Microsoft.Data.SqlClient.SqlConnection(cs);
                cn.Open();

                using var tx = cn.BeginTransaction();

                foreach (var row in rows)
                {
                    string updateEmployeeSql = @"
                UPDATE [dbo].[Employee Records]
                SET
                    [Tip Amount] = @CashTips,
                    [Credit Card Tips] = @CreditCardTips,
                    [Tips Paid Out] = @TipsPaidOut,
                    [Tip Share] = @TipShare,
                    [Altered By] = 'MSaleWebServer'
                WHERE [RecordId] = @RecordId";

                    using (var cmd = new Microsoft.Data.SqlClient.SqlCommand(updateEmployeeSql, cn, tx))
                    {
                        cmd.Parameters.AddWithValue("@RecordId", row.RecordId ?? "");
                        cmd.Parameters.AddWithValue("@CashTips", row.CashTips);
                        cmd.Parameters.AddWithValue("@CreditCardTips", row.CreditCardTips);
                        cmd.Parameters.AddWithValue("@TipsPaidOut", row.TipsPaidOut);
                        cmd.Parameters.AddWithValue("@TipShare", row.TipShare);
                        cmd.ExecuteNonQuery();
                    }

                    if (row.Shift > 0)
                    {
                        string upsertTipIncomeSql = @"
                    IF EXISTS (
                        SELECT 1
                        FROM [dbo].[Tip Income]
                        WHERE [Server Name] = @EmployeeName
                          AND TRY_CAST([Date] AS date) = TRY_CAST(@ShiftDate AS date)
                          AND ISNULL([Shift], 0) = @Shift
                    )
                    BEGIN
                        UPDATE [dbo].[Tip Income]
                        SET
                            [Declared Tips] = @CashTips,
                            [Charged Tips] = @CreditCardTips,
                            [Tips Paid Out] = @TipsPaidOut,
                            [Tip Share] = @TipShare,
                            [Gratuity] = @Gratuity
                        WHERE [Server Name] = @EmployeeName
                          AND TRY_CAST([Date] AS date) = TRY_CAST(@ShiftDate AS date)
                          AND ISNULL([Shift], 0) = @Shift
                    END
                    ELSE
                    BEGIN
                        INSERT INTO [dbo].[Tip Income]
                        (
                            [Date],
                            [Server Name],
                            [Shift],
                            [Declared Tips],
                            [Charged Tips],
                            [Tips Paid Out],
                            [Tip Share],
                            [Gratuity]
                        )
                        VALUES
                        (
                            @ShiftDate,
                            @EmployeeName,
                            @Shift,
                            @CashTips,
                            @CreditCardTips,
                            @TipsPaidOut,
                            @TipShare,
                            @Gratuity
                        )
                    END";

                        using var cmd = new Microsoft.Data.SqlClient.SqlCommand(upsertTipIncomeSql, cn, tx);
                        cmd.Parameters.AddWithValue("@EmployeeName", row.EmployeeName ?? "");
                        cmd.Parameters.AddWithValue("@ShiftDate", row.ShiftDate ?? "");
                        cmd.Parameters.AddWithValue("@Shift", row.Shift);
                        cmd.Parameters.AddWithValue("@CashTips", row.CashTips);
                        cmd.Parameters.AddWithValue("@CreditCardTips", row.CreditCardTips);
                        cmd.Parameters.AddWithValue("@TipsPaidOut", row.TipsPaidOut);
                        cmd.Parameters.AddWithValue("@TipShare", row.TipShare);
                        cmd.Parameters.AddWithValue("@Gratuity", row.Gratuity);
                        cmd.ExecuteNonQuery();
                    }
                }

                tx.Commit();

                return Json(new { success = true, message = "Tips saved." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
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
