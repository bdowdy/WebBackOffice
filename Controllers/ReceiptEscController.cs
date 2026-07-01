using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    // Editor for receipt print size (menu items + modifiers), per printer.
    //
    // Clean storage in [Receipt Print Format] (one row per printer):
    //   [Item Size] / [Modifier Size] = 'Large' | 'Small'
    //   [Big Esc]   / [Small Esc]     = comma-separated decimal ESC bytes
    //
    // The VB6 POS reads this row and wraps item / modifier lines with the
    // Big or Small ESC sequence accordingly. See ReceiptPrintFormat_Schema.sql.
    [ApiController]
    [Route("api/receipt-esc")]
    public class ReceiptEscController : ControllerBase
    {
        private const string DefaultBig = "29,33,17"; // GS ! 0x11 double height + width
        private const string DefaultSmall = "29,33,0"; // GS ! 0x00 normal

        private readonly IConfiguration _configuration;
        public ReceiptEscController(IConfiguration configuration) => _configuration = configuration;

        private string Cs => _configuration.GetConnectionString("DefaultConnection");

        public class FormatDto
        {
            public string Printer { get; set; } = "";
            public string ItemSize { get; set; } = "Small";     // 'Small' | 'Large'
            public string ModifierSize { get; set; } = "Small";
            public string BigEsc { get; set; } = DefaultBig;
            public string SmallEsc { get; set; } = DefaultSmall;
        }

        // GET /api/receipt-esc/printers
        [HttpGet("printers")]
        public IActionResult GetPrinters()
        {
            try
            {
                var list = new List<string>();
                using var cn = new SqlConnection(Cs);
                cn.Open();
                using var cmd = new SqlCommand(
                    "SELECT DISTINCT [Printer Name] FROM [dbo].[System Printers] " +
                    "WHERE [Printer Name] IS NOT NULL AND LTRIM(RTRIM([Printer Name])) NOT IN ('', 'None', 'Not Available') " +
                    "ORDER BY [Printer Name]", cn);
                using var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(r.GetString(0));
                return Ok(list);
            }
            catch (Exception ex) { return StatusCode(500, "Error: " + ex.Message); }
        }

        // GET /api/receipt-esc/format?printer=NAME
        [HttpGet("format")]
        public IActionResult GetFormat([FromQuery] string printer)
        {
            if (string.IsNullOrWhiteSpace(printer))
                return BadRequest("Printer is required.");
            try
            {
                using var cn = new SqlConnection(Cs);
                cn.Open();
                using var cmd = new SqlCommand(
                    "SELECT [Item Size], [Modifier Size], [Big Esc], [Small Esc] " +
                    "FROM [dbo].[Receipt Print Format] WHERE [Printer Name] = @p", cn);
                cmd.Parameters.Add("@p", System.Data.SqlDbType.NVarChar, 100).Value = printer;

                using var r = cmd.ExecuteReader();
                if (r.Read())
                {
                    return Ok(new FormatDto
                    {
                        Printer = printer,
                        ItemSize = Str(r, 0, "Small"),
                        ModifierSize = Str(r, 1, "Small"),
                        BigEsc = Str(r, 2, DefaultBig),
                        SmallEsc = Str(r, 3, DefaultSmall)
                    });
                }
                // No row yet -> defaults.
                return Ok(new FormatDto { Printer = printer });
            }
            catch (Exception ex) { return StatusCode(500, "Error: " + ex.Message); }
        }

        // POST /api/receipt-esc/save
        [HttpPost("save")]
        public IActionResult Save([FromBody] FormatDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.Printer))
                return BadRequest("Printer is required.");

            string item = NormSize(dto.ItemSize);
            string mod = NormSize(dto.ModifierSize);
            string big = NormCsv(dto.BigEsc, DefaultBig);
            string small = NormCsv(dto.SmallEsc, DefaultSmall);

            try
            {
                using var cn = new SqlConnection(Cs);
                cn.Open();
                using var cmd = new SqlCommand(@"
UPDATE [dbo].[Receipt Print Format]
   SET [Item Size] = @item, [Modifier Size] = @mod,
       [Big Esc] = @big, [Small Esc] = @small, [Modified Utc] = SYSUTCDATETIME()
 WHERE [Printer Name] = @p;
IF @@ROWCOUNT = 0
   INSERT INTO [dbo].[Receipt Print Format]
        ([Printer Name], [Item Size], [Modifier Size], [Big Esc], [Small Esc])
   VALUES (@p, @item, @mod, @big, @small);", cn);
                cmd.Parameters.Add("@p", System.Data.SqlDbType.NVarChar, 100).Value = dto.Printer;
                cmd.Parameters.Add("@item", System.Data.SqlDbType.NVarChar, 10).Value = item;
                cmd.Parameters.Add("@mod", System.Data.SqlDbType.NVarChar, 10).Value = mod;
                cmd.Parameters.Add("@big", System.Data.SqlDbType.NVarChar, 100).Value = big;
                cmd.Parameters.Add("@small", System.Data.SqlDbType.NVarChar, 100).Value = small;
                cmd.ExecuteNonQuery();

                return NoContent();
            }
            catch (Exception ex) { return StatusCode(500, "Error: " + ex.Message); }
        }

        // --- helpers -------------------------------------------------------

        private static string Str(SqlDataReader r, int i, string fallback)
            => r.IsDBNull(i) ? fallback : (r.GetValue(i)?.ToString() ?? fallback);

        private static string NormSize(string s)
            => string.Equals(s, "Large", StringComparison.OrdinalIgnoreCase) ? "Large" : "Small";

        // Keep only integers 0-255, comma-joined. Falls back if empty/invalid.
        private static string NormCsv(string csv, string fallback)
        {
            if (string.IsNullOrWhiteSpace(csv)) return fallback;
            var parts = csv.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var outv = new List<string>();
            foreach (var p in parts)
                if (int.TryParse(p.Trim(), out var n) && n >= 0 && n <= 255)
                    outv.Add(n.ToString());
            return outv.Count > 0 ? string.Join(",", outv) : fallback;
        }
    }
}
