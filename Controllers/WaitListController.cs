using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MSaleWebServer.Services;

namespace MSaleWebServer.Controllers
{
    public class WaitListController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly SmsService     _sms;

        public WaitListController(IConfiguration configuration, SmsService sms)
        {
            _configuration = configuration;
            _sms           = sms;
        }

        [HttpGet]
        public IActionResult Index() => View("~/Views/MSaleAdmin/WaitList.cshtml");

        // ── Get list ─────────────────────────────────────────────────────────
        // filter: "waiting" | "all"    sort: "timein" | "party" | "preference"
        [HttpGet]
        public IActionResult GetAll([FromQuery] string filter = "waiting",
                                    [FromQuery] string sort   = "timein",
                                    [FromQuery] string search = "")
        {
            var list = new List<object>();
            try
            {
                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();

                string whereSeated = filter == "all" ? "" : "AND [Seated] = 0";
                string whereName   = string.IsNullOrWhiteSpace(search)
                    ? "" : $"AND [Name] LIKE @Search";

                string orderBy = sort switch
                {
                    "party"      => "CAST([Time In] AS datetime2), [Number in Party]",
                    "preference" => "CAST([Time In] AS datetime2), [Preference]",
                    _            => "CAST([Time In] AS datetime2)"
                };

                string sql = $@"
                    SELECT
                        [ID], [Day], [Seated], [Name],
                        [Number in Party] AS PartySize,
                        [Time In]         AS TimeIn,
                        [Preference], [Quoted],
                        [Time Seated]     AS TimeSeated,
                        [Pager]           AS Cell,
                        [Group]           AS Area,
                        [TextTime]
                    FROM [Wait List]
                    WHERE [Day] = 0
                    {whereSeated}
                    {whereName}
                    ORDER BY {orderBy}";

                using var cmd = new SqlCommand(sql, cn);
                if (!string.IsNullOrWhiteSpace(search))
                    cmd.Parameters.AddWithValue("@Search", search.Trim() + "%");

                using var dr = cmd.ExecuteReader();
                var now = DateTime.Now;

                while (dr.Read())
                {
                    string timeIn  = dr["TimeIn"]?.ToString()?.Trim()  ?? "";
                    int    quoted  = Convert.ToInt32(dr["Quoted"]);
                    int    elapsed = 0;

                    if (TimeSpan.TryParse(timeIn, out var tin))
                    {
                        var arrivalToday = DateTime.Today.Add(tin);
                        elapsed = (int)(now - arrivalToday).TotalMinutes;
                        if (elapsed < 0) elapsed = 0;
                    }

                    list.Add(new {
                        id          = Convert.ToInt32(dr["ID"]),
                        seated      = Convert.ToBoolean(dr["Seated"]),
                        name        = dr["Name"]?.ToString()       ?? "",
                        partySize   = Convert.ToInt32(dr["PartySize"]),
                        timeIn,
                        elapsed,
                        preference  = dr["Preference"]?.ToString() ?? "",
                        quoted,
                        overdue     = !Convert.ToBoolean(dr["Seated"]) && quoted > 0 && elapsed >= quoted,
                        timeSeated  = dr["TimeSeated"]?.ToString() ?? "",
                        cell        = dr["Cell"]?.ToString()        ?? "",
                        area        = dr["Area"]?.ToString()        ?? "",
                        textTime    = dr["TextTime"]?.ToString()    ?? ""
                    });
                }
            }
            catch (Exception ex)
            {
                return Json(new { error = ex.Message });
            }
            return Json(list);
        }

        // ── Stats ─────────────────────────────────────────────────────────────
        [HttpGet]
        public IActionResult GetStats()
        {
            try
            {
                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();

                // Parties + guests waiting
                int parties = 0, guests = 0;
                using (var cmd = new SqlCommand(
                    "SELECT COUNT(*) AS P, ISNULL(SUM([Number in Party]),0) AS G FROM [Wait List] WHERE [Day]=0 AND [Seated]=0", cn))
                using (var dr = cmd.ExecuteReader())
                {
                    if (dr.Read()) { parties = Convert.ToInt32(dr["P"]); guests = Convert.ToInt32(dr["G"]); }
                }

                // Average wait (seated today)
                double avgWait = 0;
                using (var cmd = new SqlCommand(
                    "SELECT [Time In],[Time Seated] FROM [Wait List] WHERE [Day]=0 AND [Seated]=1 AND [Time Seated]<>''", cn))
                using (var dr = cmd.ExecuteReader())
                {
                    var mins = new List<double>();
                    while (dr.Read())
                    {
                        if (TimeSpan.TryParse(dr["Time In"]?.ToString(),     out var tIn) &&
                            TimeSpan.TryParse(dr["Time Seated"]?.ToString(), out var tSeat))
                        {
                            var diff = (tSeat - tIn).TotalMinutes;
                            if (diff > 0 && diff < 240) mins.Add(diff);
                        }
                    }
                    if (mins.Count > 0) avgWait = mins.Average();
                }

                return Json(new {
                    parties,
                    guests,
                    avgWait = avgWait > 0 ? $"{(int)avgWait} min" : "—"
                });
            }
            catch (Exception ex)
            {
                return Json(new { error = ex.Message });
            }
        }

        // ── Add party ─────────────────────────────────────────────────────────
        [HttpPost]
        public IActionResult Add([FromBody] WaitListDto dto)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(dto.Name))
                    return Json(new { success = false, message = "Name is required." });

                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();

                string timeIn = string.IsNullOrWhiteSpace(dto.TimeIn)
                    ? DateTime.Now.ToString("HH:mm")
                    : dto.TimeIn;

                using var cmd = new SqlCommand(@"
                    INSERT INTO [Wait List]
                        ([Day],[Date],[Seated],[Name],[Number in Party],
                         [Time In],[Preference],[Quoted],[Time Seated],[Pager],[Group],[TextTime])
                    VALUES
                        (0, @Date, 0, @Name, @PartySize,
                         @TimeIn, @Pref, @Quoted, '', @Cell, @Area, '')", cn);

                cmd.Parameters.AddWithValue("@Date",      DateTime.Today.ToShortDateString());
                cmd.Parameters.AddWithValue("@Name",      dto.Name.Trim());
                cmd.Parameters.AddWithValue("@PartySize", dto.PartySize);
                cmd.Parameters.AddWithValue("@TimeIn",    timeIn);
                cmd.Parameters.AddWithValue("@Pref",      dto.Preference ?? "");
                cmd.Parameters.AddWithValue("@Quoted",    dto.Quoted);
                cmd.Parameters.AddWithValue("@Cell",      dto.Cell ?? "");
                cmd.Parameters.AddWithValue("@Area",      dto.Area ?? "ALL");
                cmd.ExecuteNonQuery();

                // Send confirmation text if cell number provided
                if (!string.IsNullOrWhiteSpace(dto.Cell))
                {
                    string waitMsg = dto.Quoted > 0
                        ? $"Hi {dto.Name.Trim()}, you're on our wait list! Estimated wait: {dto.Quoted} min. We'll text you when your table is ready."
                        : $"Hi {dto.Name.Trim()}, you're on our wait list! We'll text you when your table is ready.";
                    _sms.Send(dto.Cell, waitMsg);
                }

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // ── Mark seated ───────────────────────────────────────────────────────
        [HttpPost]
        public IActionResult SetSeated([FromBody] SetSeatedDto dto)
        {
            try
            {
                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();

                string timeSeated = dto.Seated ? DateTime.Now.ToString("HH:mm") : "";
                using var cmd = new SqlCommand(
                    "UPDATE [Wait List] SET [Seated]=@S, [Time Seated]=@TS WHERE [ID]=@ID", cn);
                cmd.Parameters.AddWithValue("@S",  dto.Seated ? 1 : 0);
                cmd.Parameters.AddWithValue("@TS", timeSeated);
                cmd.Parameters.AddWithValue("@ID", dto.Id);
                cmd.ExecuteNonQuery();

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // ── Send table-ready text + stamp TextTime ────────────────────────────
        [HttpPost]
        public IActionResult SendText([FromBody] SendTextDto dto)
        {
            try
            {
                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();

                // Fetch name + cell for the message
                string name = "", cell = "";
                using (var sel = new SqlCommand(
                    "SELECT [Name],[Pager] FROM [Wait List] WHERE [ID]=@ID", cn))
                {
                    sel.Parameters.AddWithValue("@ID", dto.Id);
                    using var dr = sel.ExecuteReader();
                    if (dr.Read())
                    {
                        name = dr["Name"]?.ToString()  ?? "";
                        cell = dr["Pager"]?.ToString() ?? "";
                    }
                }

                // Send the text
                if (!string.IsNullOrWhiteSpace(cell))
                {
                    string msg = $"Hi {name.Trim()}, your table is ready! Please see the host stand. We look forward to serving you.";
                    _sms.Send(cell, msg);
                }

                // Stamp the time regardless (so staff can see it was triggered)
                string ts = DateTime.Now.ToString("HH:mm");
                using var upd = new SqlCommand(
                    "UPDATE [Wait List] SET [TextTime]=@T WHERE [ID]=@ID", cn);
                upd.Parameters.AddWithValue("@T",  ts);
                upd.Parameters.AddWithValue("@ID", dto.Id);
                upd.ExecuteNonQuery();

                return Json(new { success = true, textTime = ts });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // ── Clear (archive seated rows) ───────────────────────────────────────
        [HttpPost]
        public IActionResult ClearList()
        {
            try
            {
                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();
                int dateNo = int.Parse(DateTime.Today.ToString("yyyyMMdd"));
                using var cmd = new SqlCommand(
                    "UPDATE [Wait List] SET [Day]=@D WHERE [Day]=0 AND [Seated]=1", cn);
                cmd.Parameters.AddWithValue("@D", dateNo);
                int rows = cmd.ExecuteNonQuery();
                return Json(new { success = true, archived = rows });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // ── Pull pending reservations into wait list ───────────────────────────
        // Matches reservations for today where reservation time is within the next 60 min
        [HttpPost]
        public IActionResult PullReservations()
        {
            int added = 0;
            try
            {
                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();

                string today   = DateTime.Today.ToString("yyyy-MM-dd");
                string cutoff  = DateTime.Now.AddMinutes(60).ToString("HH:mm");

                // Read reservations due within the next hour
                var pending = new List<(string name, int size, string time, string pref, string phone)>();
                using (var cmd = new SqlCommand(@"
                    SELECT [Party Name], [Party Size], [Reservation Time], [Notes], [Phone]
                    FROM   [Reservations]
                    WHERE  CONVERT(date, [Reservation Date]) = @Today
                      AND  [Reservation Time] <= @Cutoff
                      AND  ISNULL([Party Name],'') <> ''", cn))
                {
                    cmd.Parameters.AddWithValue("@Today",  today);
                    cmd.Parameters.AddWithValue("@Cutoff", cutoff);
                    using var dr = cmd.ExecuteReader();
                    while (dr.Read())
                        pending.Add((
                            dr["Party Name"]?.ToString()        ?? "",
                            Convert.ToInt32(dr["Party Size"]),
                            dr["Reservation Time"]?.ToString()  ?? "",
                            dr["Notes"]?.ToString()             ?? "",
                            dr["Phone"]?.ToString()             ?? ""
                        ));
                }

                foreach (var r in pending)
                {
                    // Skip if already in wait list today
                    using var chk = new SqlCommand(@"
                        SELECT COUNT(*) FROM [Wait List]
                        WHERE [Day]=0 AND [Name]=@N AND [Pager]=@P AND [Time In]=@T", cn);
                    chk.Parameters.AddWithValue("@N", r.name);
                    chk.Parameters.AddWithValue("@P", r.phone);
                    chk.Parameters.AddWithValue("@T", r.time);
                    if (Convert.ToInt32(chk.ExecuteScalar()) > 0) continue;

                    using var ins = new SqlCommand(@"
                        INSERT INTO [Wait List]
                            ([Day],[Date],[Seated],[Name],[Number in Party],
                             [Time In],[Preference],[Quoted],[Time Seated],[Pager],[Group],[TextTime])
                        VALUES (0,@Date,0,@Name,@PS,@TI,@Pref,0,'',@Cell,'ALL','')", cn);
                    ins.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());
                    ins.Parameters.AddWithValue("@Name", r.name);
                    ins.Parameters.AddWithValue("@PS",   r.size);
                    ins.Parameters.AddWithValue("@TI",   r.time);
                    ins.Parameters.AddWithValue("@Pref", r.pref);
                    ins.Parameters.AddWithValue("@Cell", r.phone);
                    ins.ExecuteNonQuery();
                    added++;
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
            return Json(new { success = true, added });
        }

        // ── Get preference options from existing data ──────────────────────────
        [HttpGet]
        public IActionResult GetPreferences()
        {
            var list = new List<string> { "Any", "Booth", "Table", "Bar", "High Top", "Patio", "Outside" };
            return Json(list);
        }
    }

    public class WaitListDto
    {
        public string Name       { get; set; } = "";
        public int    PartySize  { get; set; } = 1;
        public string TimeIn     { get; set; } = "";
        public string Preference { get; set; } = "";
        public int    Quoted     { get; set; } = 0;
        public string Cell       { get; set; } = "";
        public string Area       { get; set; } = "ALL";
    }

    public class SetSeatedDto
    {
        public int  Id     { get; set; }
        public bool Seated { get; set; }
    }

    public class SendTextDto
    {
        public int Id { get; set; }
    }
}
