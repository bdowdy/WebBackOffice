using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MSaleWebServer.Services;

namespace MSaleWebServer.Controllers
{
    public class ReservationController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly SmsService     _sms;

        public ReservationController(IConfiguration configuration, SmsService sms)
        {
            _configuration = configuration;
            _sms           = sms;
        }

        [HttpGet]
        public IActionResult Index() => View("~/Views/MSaleAdmin/Reservation.cshtml");

        // ── Public booking page (/Book) ───────────────────────────────────────
        [HttpGet("/Book")]
        public IActionResult Book() => View("~/Views/MSaleAdmin/Book.cshtml");

        [HttpPost("/Book/Submit")]
        public IActionResult BookSubmit([FromBody] ReservationDto dto)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(dto.PartyName))
                    return Json(new { success = false, message = "Please enter your name." });
                if (string.IsNullOrWhiteSpace(dto.ReservationDate))
                    return Json(new { success = false, message = "Please select a date." });
                if (string.IsNullOrWhiteSpace(dto.ReservationTime))
                    return Json(new { success = false, message = "Please select a time." });

                dto.Status = "Pending";

                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();
                using var cmd = new SqlCommand(@"
                    INSERT INTO [Reservations]
                        ([Party Name], [Party Size], [Reservation Date], [Reservation Time],
                         [Phone], [Notes], [Table Number], [Status])
                    VALUES (@PN, @PS, @RD, @RT, @PH, @NO, '', @ST)", cn);
                AddParams(cmd, dto);
                cmd.ExecuteNonQuery();

                // Send booking confirmation text
                if (!string.IsNullOrWhiteSpace(dto.Phone))
                {
                    var d = DateTime.TryParse(dto.ReservationDate, out var dt)
                        ? dt.ToString("dddd, MMMM d")
                        : dto.ReservationDate;

                    var t = TimeSpan.TryParse(dto.ReservationTime, out var ts)
                        ? DateTime.Today.Add(ts).ToString("h:mm tt")
                        : dto.ReservationTime;

                    string guests = dto.PartySize == 1 ? "1 guest" : $"{dto.PartySize} guests";
                    string msg    = $"Hi {dto.PartyName.Trim()}, your reservation is confirmed! {guests} on {d} at {t}. Reply STOP to opt out.";
                    _sms.Send(dto.Phone, msg);
                }

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Something went wrong. Please call us to book. (" + ex.Message + ")" });
            }
        }

        // GET /Reservation/GetAll?date=2026-06-17&status=Pending
        [HttpGet]
        public IActionResult GetAll([FromQuery] string? date, [FromQuery] string? status)
        {
            var list = new List<object>();
            try
            {
                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();

                string sql = @"
                    SELECT
                        ISNULL([Reservation ID], 0)                                AS ReservationId,
                        ISNULL([Party Name], '')                                    AS PartyName,
                        ISNULL([Party Size], 0)                                     AS PartySize,
                        ISNULL(CONVERT(varchar(10), [Reservation Date], 120), '')   AS ReservationDate,
                        ISNULL([Reservation Time], '')                              AS ReservationTime,
                        ISNULL([Phone], '')                                         AS Phone,
                        ISNULL([Notes], '')                                         AS Notes,
                        ISNULL([Table Number], '')                                  AS TableNumber,
                        ISNULL([Status], 'Pending')                                 AS Status
                    FROM [Reservations]
                    WHERE (@Date  IS NULL OR CONVERT(date, [Reservation Date]) = @Date)
                      AND (@Status IS NULL OR [Status] = @Status)
                    ORDER BY [Reservation Date], [Reservation Time], [Party Name]";

                using var cmd = new SqlCommand(sql, cn);
                cmd.Parameters.AddWithValue("@Date",   string.IsNullOrWhiteSpace(date)   ? (object)DBNull.Value : date);
                cmd.Parameters.AddWithValue("@Status", string.IsNullOrWhiteSpace(status) ? (object)DBNull.Value : status);

                using var dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    list.Add(new {
                        reservationId   = Convert.ToInt32(dr["ReservationId"]),
                        partyName       = dr["PartyName"].ToString(),
                        partySize       = Convert.ToInt32(dr["PartySize"]),
                        reservationDate = dr["ReservationDate"].ToString(),
                        reservationTime = dr["ReservationTime"].ToString(),
                        phone           = dr["Phone"].ToString(),
                        notes           = dr["Notes"].ToString(),
                        tableNumber     = dr["TableNumber"].ToString(),
                        status          = dr["Status"].ToString()
                    });
                }
            }
            catch (Exception ex)
            {
                return Json(new { error = ex.Message });
            }
            return Json(list);
        }

        [HttpPost]
        public IActionResult Save([FromBody] ReservationDto dto)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(dto.PartyName))
                    return Json(new { success = false, message = "Party name is required." });

                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();

                if (dto.ReservationId == 0)
                {
                    using var cmd = new SqlCommand(@"
                        INSERT INTO [Reservations]
                            ([Party Name], [Party Size], [Reservation Date], [Reservation Time],
                             [Phone], [Notes], [Table Number], [Status])
                        VALUES (@PN, @PS, @RD, @RT, @PH, @NO, @TN, @ST)", cn);
                    AddParams(cmd, dto);
                    cmd.ExecuteNonQuery();
                }
                else
                {
                    using var cmd = new SqlCommand(@"
                        UPDATE [Reservations] SET
                            [Party Name]       = @PN,
                            [Party Size]       = @PS,
                            [Reservation Date] = @RD,
                            [Reservation Time] = @RT,
                            [Phone]            = @PH,
                            [Notes]            = @NO,
                            [Table Number]     = @TN,
                            [Status]           = @ST
                        WHERE [Reservation ID] = @ID", cn);
                    AddParams(cmd, dto);
                    cmd.Parameters.AddWithValue("@ID", dto.ReservationId);
                    cmd.ExecuteNonQuery();
                }
                return Json(new { success = true, message = "Reservation saved." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public IActionResult Delete([FromBody] DeleteReservationDto dto)
        {
            try
            {
                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();
                using var cmd = new SqlCommand(
                    "DELETE FROM [Reservations] WHERE [Reservation ID] = @ID", cn);
                cmd.Parameters.AddWithValue("@ID", dto.ReservationId);
                cmd.ExecuteNonQuery();
                return Json(new { success = true, message = "Reservation deleted." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // Diagnostic — returns column names so you can verify the table schema
        [HttpGet]
        public IActionResult GetColumns()
        {
            var cols = new List<object>();
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();
            using var cmd = new SqlCommand(@"
                SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'Reservations'
                ORDER BY ORDINAL_POSITION", cn);
            using var dr = cmd.ExecuteReader();
            while (dr.Read())
                cols.Add(new { column = dr["COLUMN_NAME"].ToString(), type = dr["DATA_TYPE"].ToString() });
            return Json(cols);
        }

        // Returns table names from the Layout table for the table-assignment dropdown
        [HttpGet]
        public IActionResult GetTables()
        {
            var list = new List<string>();
            try
            {
                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();
                using var cmd = new SqlCommand(@"
                    SELECT DISTINCT ISNULL([Table Name], '') AS [Table Name]
                    FROM [Layout]
                    WHERE LTRIM(RTRIM(ISNULL([Table Name], ''))) <> ''
                    ORDER BY [Table Name]", cn);
                using var dr = cmd.ExecuteReader();
                while (dr.Read()) list.Add(dr["Table Name"].ToString() ?? "");
            }
            catch { /* Layout table may not exist — table dropdown stays empty */ }
            return Json(list);
        }

        private static void AddParams(SqlCommand cmd, ReservationDto dto)
        {
            cmd.Parameters.AddWithValue("@PN", (object?)dto.PartyName   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@PS", dto.PartySize);
            cmd.Parameters.AddWithValue("@RD", string.IsNullOrWhiteSpace(dto.ReservationDate) ? (object)DBNull.Value : dto.ReservationDate);
            cmd.Parameters.AddWithValue("@RT", dto.ReservationTime ?? "");
            cmd.Parameters.AddWithValue("@PH", dto.Phone           ?? "");
            cmd.Parameters.AddWithValue("@NO", dto.Notes           ?? "");
            cmd.Parameters.AddWithValue("@TN", dto.TableNumber     ?? "");
            cmd.Parameters.AddWithValue("@ST", dto.Status          ?? "Pending");
        }
    }

    public class ReservationDto
    {
        public int    ReservationId   { get; set; }
        public string PartyName       { get; set; } = "";
        public int    PartySize       { get; set; } = 1;
        public string ReservationDate { get; set; } = "";
        public string ReservationTime { get; set; } = "";
        public string Phone           { get; set; } = "";
        public string Notes           { get; set; } = "";
        public string TableNumber     { get; set; } = "";
        public string Status          { get; set; } = "Pending";
    }

    public class DeleteReservationDto
    {
        public int ReservationId { get; set; }
    }
}
