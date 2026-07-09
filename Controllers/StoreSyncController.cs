using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    /// <summary>
    /// Endpoints used by the store's StoreSyncAgent (X-Api-Key authenticated).
    /// - POST api/storesync/replace       : replace-recent-window upload (sales, time records)
    /// - GET  api/storesync/timerecord-edits : pending punch edits made in the web back office
    /// - POST api/storesync/timerecord-edits/ack : mark edits as applied at the store
    /// </summary>
    [ApiController]
    [Route("api/storesync")]
    public class StoreSyncController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        public StoreSyncController(IConfiguration configuration) => _configuration = configuration;

        // Legacy tables have no PKs, so sync works by replacing whole recent
        // key-windows atomically. Only these tables may be touched.
        private static readonly Dictionary<string, (string KeyColumn, bool KeyIsDate)> Allowed = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Closed Checks"]    = ("Order Date", false),
            ["Check Detail"]     = ("Check Date", false),
            ["CheckPayments"]    = ("Date",       true),
            ["Employee Records"] = ("Shift Date", false),
        };

        public class ReplaceRequest
        {
            public string Table { get; set; } = "";
            public List<string> KeyValues { get; set; } = new();
            public List<Dictionary<string, JsonElement>> Rows { get; set; } = new();
        }

        [HttpPost("replace")]
        public IActionResult Replace([FromBody] ReplaceRequest req)
        {
            if (!Allowed.TryGetValue(req.Table, out var meta))
                return BadRequest(new { message = $"Table '{req.Table}' is not syncable." });
            if (req.KeyValues.Count == 0)
                return Ok(new { deleted = 0, inserted = 0 });

            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();

            // valid column names for this table
            var cols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var cc = new SqlCommand("SELECT name FROM sys.columns WHERE object_id = OBJECT_ID(@t)", cn))
            {
                cc.Parameters.AddWithValue("@t", $"[dbo].[{req.Table}]");
                using var rd = cc.ExecuteReader();
                while (rd.Read()) cols.Add(rd.GetString(0));
            }

            using var tx = cn.BeginTransaction();
            try
            {
                // 1. delete the window
                var keyParams = string.Join(",", req.KeyValues.Select((_, i) => $"@k{i}"));
                var predicate = meta.KeyIsDate
                    ? $"CONVERT(date, [{meta.KeyColumn}]) IN ({keyParams})"
                    : $"[{meta.KeyColumn}] IN ({keyParams})";
                int deleted;
                using (var del = new SqlCommand($"DELETE FROM [dbo].[{req.Table}] WHERE {predicate}", cn, tx))
                {
                    for (int i = 0; i < req.KeyValues.Count; i++)
                        del.Parameters.AddWithValue($"@k{i}", req.KeyValues[i]);
                    deleted = del.ExecuteNonQuery();
                }

                // 2. insert replacement rows
                int inserted = 0;
                foreach (var row in req.Rows)
                {
                    var rowCols = row.Keys.Where(k => cols.Contains(k)).ToList();
                    if (rowCols.Count == 0) continue;
                    var colList = string.Join(",", rowCols.Select(c => $"[{c}]"));
                    var parList = string.Join(",", rowCols.Select((_, i) => $"@p{i}"));
                    using var ins = new SqlCommand($"INSERT INTO [dbo].[{req.Table}] ({colList}) VALUES ({parList})", cn, tx);
                    for (int i = 0; i < rowCols.Count; i++)
                        ins.Parameters.AddWithValue($"@p{i}", ToDbValue(row[rowCols[i]]));
                    inserted += ins.ExecuteNonQuery();
                }

                tx.Commit();
                return Ok(new { deleted, inserted });
            }
            catch
            {
                tx.Rollback();
                throw;
            }
        }

        private static object ToDbValue(JsonElement e) => e.ValueKind switch
        {
            JsonValueKind.String => e.GetString()!,
            JsonValueKind.Number => e.GetDecimal(),
            JsonValueKind.True   => true,
            JsonValueKind.False  => false,
            _                    => DBNull.Value
        };

        // ── Time-record edit queue (written by TimeRecordsEditor, read by the agent) ──

        internal static void EnsureQueueTable(SqlConnection cn)
        {
            const string sql = @"
IF OBJECT_ID('[dbo].[TimeRecordEdits]') IS NULL
CREATE TABLE [dbo].[TimeRecordEdits](
    [Id] int IDENTITY(1,1) PRIMARY KEY,
    [Action] nvarchar(10) NOT NULL,
    [RecordId] nvarchar(100) NOT NULL,
    [Payload] nvarchar(max) NULL,
    [CreatedUtc] datetime NOT NULL DEFAULT GETUTCDATE(),
    [AppliedUtc] datetime NULL)";
            using var cmd = new SqlCommand(sql, cn);
            cmd.ExecuteNonQuery();
        }

        internal static void QueueEdit(SqlConnection cn, string action, string recordId, string payloadJson)
        {
            EnsureQueueTable(cn);
            using var cmd = new SqlCommand(
                "INSERT INTO [dbo].[TimeRecordEdits] ([Action],[RecordId],[Payload]) VALUES (@a,@r,@p)", cn);
            cmd.Parameters.AddWithValue("@a", action);
            cmd.Parameters.AddWithValue("@r", recordId);
            cmd.Parameters.AddWithValue("@p", (object?)payloadJson ?? DBNull.Value);
            cmd.ExecuteNonQuery();
        }

        [HttpGet("timerecord-edits")]
        public IActionResult PendingEdits()
        {
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();
            EnsureQueueTable(cn);
            var list = new List<object>();
            using var cmd = new SqlCommand(
                "SELECT [Id],[Action],[RecordId],[Payload] FROM [dbo].[TimeRecordEdits] WHERE [AppliedUtc] IS NULL ORDER BY [Id]", cn);
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
                list.Add(new
                {
                    id = rd.GetInt32(0),
                    action = rd.GetString(1),
                    recordId = rd.GetString(2),
                    payload = rd.IsDBNull(3) ? null : rd.GetString(3)
                });
            return Ok(list);
        }

        public class AckRequest { public List<int> Ids { get; set; } = new(); }

        [HttpPost("timerecord-edits/ack")]
        public IActionResult AckEdits([FromBody] AckRequest req)
        {
            if (req.Ids.Count == 0) return Ok(new { acked = 0 });
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();
            var pars = string.Join(",", req.Ids.Select((_, i) => $"@i{i}"));
            using var cmd = new SqlCommand(
                $"UPDATE [dbo].[TimeRecordEdits] SET [AppliedUtc] = GETUTCDATE() WHERE [Id] IN ({pars})", cn);
            for (int i = 0; i < req.Ids.Count; i++)
                cmd.Parameters.AddWithValue($"@i{i}", req.Ids[i]);
            var n = cmd.ExecuteNonQuery();
            return Ok(new { acked = n });
        }
    }
}
