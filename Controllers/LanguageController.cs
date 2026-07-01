using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/language")]
    public class LanguageController : ControllerBase
    {
        private const string LanguageOnSetting = "Language On";
        private readonly IConfiguration _configuration;

        public LanguageController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        private string Cs => _configuration.GetConnectionString("DefaultConnection");

        // Quote an identifier from schema metadata safely.
        private static string Q(string id) => "[" + id.Replace("]", "]]") + "]";

        // The migration to SQL Server renamed the old Access "Return Text"
        // column, so resolve the real column names from the table schema
        // instead of hard-coding them.
        private static (string engCol, string transCol) ResolveColumns(SqlConnection cn)
        {
            var cols = new List<string>();
            using (var cmd = new SqlCommand(
                "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'Language'", cn))
            using (var r = cmd.ExecuteReader())
            {
                while (r.Read()) cols.Add(r.GetString(0));
            }

            if (cols.Count == 0)
                throw new Exception("Table [dbo].[Language] was not found (or has no columns).");

            string eng = cols.FirstOrDefault(c => c.Equals("English", StringComparison.OrdinalIgnoreCase)) ?? cols[0];

            var others = cols.Where(c => !c.Equals(eng, StringComparison.OrdinalIgnoreCase)).ToList();
            string trans =
                others.FirstOrDefault(c => c.IndexOf("return", StringComparison.OrdinalIgnoreCase) >= 0
                                        || c.IndexOf("trans", StringComparison.OrdinalIgnoreCase) >= 0
                                        || c.IndexOf("text", StringComparison.OrdinalIgnoreCase) >= 0)
                ?? others.FirstOrDefault();

            if (trans == null)
                throw new Exception("No translation column found in [dbo].[Language]. Columns present: " + string.Join(", ", cols));

            return (eng, trans);
        }

        // GET /api/language
        [HttpGet]
        public IActionResult GetTranslations()
        {
            try
            {
                var list = new List<LanguageEntryDto>();

                using var cn = new SqlConnection(Cs);
                cn.Open();
                var (engCol, transCol) = ResolveColumns(cn);

                string sql = $"SELECT {Q(engCol)}, ISNULL({Q(transCol)}, N'') AS TransText " +
                             $"FROM [dbo].[Language] ORDER BY {Q(engCol)} ASC";

                using var cmd = new SqlCommand(sql, cn);
                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    list.Add(new LanguageEntryDto
                    {
                        English = rdr[0]?.ToString() ?? "",
                        ReturnText = rdr[1] == DBNull.Value ? "" : rdr[1]?.ToString() ?? ""
                    });
                }

                return Ok(list);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Error: " + ex.Message);
            }
        }

        // PUT /api/language  { english, returnText }
        [HttpPut]
        public IActionResult Upsert([FromBody] LanguageEntryDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.English))
                return BadRequest("English phrase is required.");

            try
            {
                using var cn = new SqlConnection(Cs);
                cn.Open();
                var (engCol, transCol) = ResolveColumns(cn);

                // Update-then-insert (works with or without a unique index).
                string sql =
                    $"UPDATE [dbo].[Language] SET {Q(transCol)} = @Trans WHERE {Q(engCol)} = @Eng; " +
                    $"IF @@ROWCOUNT = 0 INSERT INTO [dbo].[Language] ({Q(engCol)}, {Q(transCol)}) VALUES (@Eng, @Trans);";

                using var tx = cn.BeginTransaction();
                using (var cmd = new SqlCommand(sql, cn, tx))
                {
                    cmd.Parameters.Add("@Eng", System.Data.SqlDbType.NVarChar, 255).Value = dto.English.Trim();
                    cmd.Parameters.Add("@Trans", System.Data.SqlDbType.NVarChar, 255).Value = (dto.ReturnText ?? "").Trim();
                    cmd.ExecuteNonQuery();
                }
                tx.Commit();

                return NoContent();
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Error: " + ex.Message);
            }
        }

        // DELETE /api/language?english=...
        [HttpDelete]
        public IActionResult Delete([FromQuery] string english)
        {
            if (string.IsNullOrWhiteSpace(english))
                return BadRequest("English phrase is required.");

            try
            {
                using var cn = new SqlConnection(Cs);
                cn.Open();
                var (engCol, _) = ResolveColumns(cn);

                using var cmd = new SqlCommand($"DELETE FROM [dbo].[Language] WHERE {Q(engCol)} = @Eng", cn);
                cmd.Parameters.Add("@Eng", System.Data.SqlDbType.NVarChar, 255).Value = english;
                cmd.ExecuteNonQuery();

                return NoContent();
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Error: " + ex.Message);
            }
        }

        // GET /api/language/setting/language-on
        [HttpGet("setting/language-on")]
        public IActionResult GetLanguageOn()
        {
            try
            {
                using var cn = new SqlConnection(Cs);
                cn.Open();

                using var cmd = new SqlCommand(
                    "SELECT [SettingValue] FROM [dbo].[AppSetting] WHERE [SettingName] = @Name", cn);
                cmd.Parameters.Add("@Name", System.Data.SqlDbType.NVarChar, 100).Value = LanguageOnSetting;

                var val = cmd.ExecuteScalar();
                string raw = val == null || val == DBNull.Value ? "" : val.ToString() ?? "";
                bool on = raw == "1" || string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase);

                return Ok(new LanguageFlagDto { Value = on });
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Error: " + ex.Message);
            }
        }

        // PUT /api/language/setting/language-on  { value: true|false }
        [HttpPut("setting/language-on")]
        public IActionResult SetLanguageOn([FromBody] LanguageFlagDto dto)
        {
            try
            {
                using var cn = new SqlConnection(Cs);
                cn.Open();

                const string sql = @"
UPDATE [dbo].[AppSetting] SET [SettingValue] = @Value, [ModifiedUtc] = SYSUTCDATETIME() WHERE [SettingName] = @Name;
IF @@ROWCOUNT = 0 INSERT INTO [dbo].[AppSetting] ([SettingName], [SettingValue]) VALUES (@Name, @Value);";

                using var tx = cn.BeginTransaction();
                using (var cmd = new SqlCommand(sql, cn, tx))
                {
                    cmd.Parameters.Add("@Name", System.Data.SqlDbType.NVarChar, 100).Value = LanguageOnSetting;
                    cmd.Parameters.Add("@Value", System.Data.SqlDbType.NVarChar, 400).Value = dto != null && dto.Value ? "1" : "0";
                    cmd.ExecuteNonQuery();
                }
                tx.Commit();

                return NoContent();
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Error: " + ex.Message);
            }
        }
    }

    public class LanguageEntryDto
    {
        public string English { get; set; } = "";
        public string ReturnText { get; set; } = "";
    }

    public class LanguageFlagDto
    {
        public bool Value { get; set; }
    }
}
