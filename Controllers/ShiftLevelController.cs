using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    public class ShiftLevelController : Controller
    {
        private readonly IConfiguration _configuration;
        public ShiftLevelController(IConfiguration configuration) => _configuration = configuration;

        [HttpGet]
        public IActionResult Index() => View("~/Views/MSaleAdmin/ShiftLevel.cshtml");

        [HttpGet]
        public IActionResult GetAll()
        {
            var list = new List<object>();
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();
            using var cmd = new SqlCommand(@"
                SELECT [Key Name], ISNULL([Category Name],'') AS [Category Name]
                FROM [Shift Keys]
                ORDER BY [Key Name], [Category Name]", cn);
            using var dr = cmd.ExecuteReader();
            while (dr.Read())
                list.Add(new
                {
                    keyName = dr["Key Name"]?.ToString() ?? "",
                    categoryName = dr["Category Name"]?.ToString() ?? ""
                });
            return Json(list);
        }

        [HttpGet]
        public IActionResult Get(string keyName, string categoryName)
        {
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();
            using var cmd = new SqlCommand(@"
                SELECT TOP 1
                    ISNULL([Key Name],'')       AS [Key Name],
                    ISNULL([Shift Level],1)     AS [Shift Level],
                    ISNULL([Category Name],'')  AS [Category Name],
                    ISNULL([Event 1],'')        AS [Event 1],
                    ISNULL([Event 2],'')        AS [Event 2],
                    ISNULL([Event 3],'')        AS [Event 3],
                    ISNULL([Event 4],'')        AS [Event 4],
                    ISNULL([Event 5],'')        AS [Event 5],
                    ISNULL([Description],'')    AS [Description],
                    ISNULL([PLU Prefix],0)      AS [PLU Prefix]
                FROM [Shift Keys]
                WHERE [Key Name] = @K AND [Category Name] = @C", cn);
            cmd.Parameters.AddWithValue("@K", keyName ?? "");
            cmd.Parameters.AddWithValue("@C", categoryName ?? "");
            using var dr = cmd.ExecuteReader();
            if (!dr.Read())
                return Json(new { success = false, message = "Record not found." });
            return Json(new
            {
                success = true,
                data = new
                {
                    keyName      = dr["Key Name"].ToString(),
                    shiftLevel   = Convert.ToInt32(dr["Shift Level"]),
                    categoryName = dr["Category Name"].ToString(),
                    event1       = dr["Event 1"].ToString(),
                    event2       = dr["Event 2"].ToString(),
                    event3       = dr["Event 3"].ToString(),
                    event4       = dr["Event 4"].ToString(),
                    event5       = dr["Event 5"].ToString(),
                    description  = dr["Description"].ToString(),
                    pluPrefix    = Convert.ToInt32(dr["PLU Prefix"])
                }
            });
        }

        [HttpGet]
        public IActionResult GetCategories()
        {
            var list = new List<string>();
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();

            using var cmd1 = new SqlCommand(@"
                SELECT [Category Name] FROM [Category Information]
                WHERE [Category Name] IS NOT NULL AND LTRIM(RTRIM([Category Name])) <> ''
                ORDER BY [Category Name]", cn);
            using var dr1 = cmd1.ExecuteReader();
            while (dr1.Read()) list.Add(dr1["Category Name"].ToString() ?? "");
            dr1.Close();

            using var cmd2 = new SqlCommand(@"
                SELECT [Sub Name] FROM [Sub Categories]
                WHERE [Sub Name] IS NOT NULL AND LTRIM(RTRIM([Sub Name])) <> ''
                ORDER BY [Sub Name]", cn);
            using var dr2 = cmd2.ExecuteReader();
            while (dr2.Read()) list.Add(dr2["Sub Name"].ToString() ?? "");

            return Json(list.Distinct().OrderBy(x => x).ToList());
        }

        [HttpGet]
        public IActionResult GetTimedEvents()
        {
            var list = new List<string>();
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();
            using var cmd = new SqlCommand(@"
                SELECT [Event Name] FROM [Timed Pricing]
                WHERE [Event Name] IS NOT NULL AND LTRIM(RTRIM([Event Name])) <> ''
                ORDER BY [Event Name]", cn);
            using var dr = cmd.ExecuteReader();
            while (dr.Read()) list.Add(dr["Event Name"].ToString() ?? "");
            return Json(list);
        }

        [HttpPost]
        public IActionResult Save([FromBody] ShiftKeyDto model)
        {
            try
            {
                if (model == null || string.IsNullOrWhiteSpace(model.KeyName))
                    return Json(new { success = false, message = "Key Name is required." });

                if (model.PluPrefix > 32767)
                    return Json(new { success = false, message = "PLU Prefix must be 32767 or less." });

                // Auto-set Description if blank
                string desc = model.Description?.Trim() ?? "";
                if (string.IsNullOrEmpty(desc))
                {
                    desc = model.KeyName.TrimEnd('~', '!');
                }

                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();

                if (model.IsNew)
                {
                    using var chk = new SqlCommand(@"
                        SELECT COUNT(*) FROM [Shift Keys]
                        WHERE [Key Name] = @K AND [Category Name] = @C", cn);
                    chk.Parameters.AddWithValue("@K", model.KeyName.Trim());
                    chk.Parameters.AddWithValue("@C", model.CategoryName ?? "");
                    if (Convert.ToInt32(chk.ExecuteScalar()) > 0)
                        return Json(new { success = false, message = "A shift key with that name and category already exists." });

                    using var ins = new SqlCommand(@"
                        INSERT INTO [Shift Keys]
                            ([Key Name],[Shift Level],[Category Name],
                             [Event 1],[Event 2],[Event 3],[Event 4],[Event 5],
                             [Description],[PLU Prefix],[Button Assigned])
                        VALUES
                            (@K,@SL,@Cat,@E1,@E2,@E3,@E4,@E5,@Desc,@PLU,-1)", cn);
                    AddParams(ins, model, desc);
                    ins.ExecuteNonQuery();
                }
                else
                {
                    using var upd = new SqlCommand(@"
                        UPDATE [Shift Keys] SET
                            [Key Name]      = @K,
                            [Shift Level]   = @SL,
                            [Category Name] = @Cat,
                            [Event 1]       = @E1,
                            [Event 2]       = @E2,
                            [Event 3]       = @E3,
                            [Event 4]       = @E4,
                            [Event 5]       = @E5,
                            [Description]   = @Desc,
                            [PLU Prefix]    = @PLU
                        WHERE [Key Name] = @OrigK AND [Category Name] = @OrigC", cn);
                    AddParams(upd, model, desc);
                    upd.Parameters.AddWithValue("@OrigK", model.OriginalKeyName?.Trim() ?? "");
                    upd.Parameters.AddWithValue("@OrigC", model.OriginalCategoryName ?? "");
                    int rows = upd.ExecuteNonQuery();
                    if (rows == 0)
                        return Json(new { success = false, message = "Record not found — save failed." });
                }

                return Json(new { success = true, message = "Saved." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public IActionResult Delete([FromBody] ShiftKeyDeleteDto model)
        {
            try
            {
                if (model == null || string.IsNullOrWhiteSpace(model.KeyName))
                    return Json(new { success = false, message = "Key Name required." });

                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();
                using var cmd = new SqlCommand(@"
                    DELETE FROM [Shift Keys]
                    WHERE [Key Name] = @K AND [Category Name] = @C", cn);
                cmd.Parameters.AddWithValue("@K", model.KeyName.Trim());
                cmd.Parameters.AddWithValue("@C", model.CategoryName ?? "");
                cmd.ExecuteNonQuery();

                return Json(new { success = true, message = "Deleted." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        private static void AddParams(SqlCommand cmd, ShiftKeyDto m, string desc)
        {
            cmd.Parameters.AddWithValue("@K",    m.KeyName.Trim());
            cmd.Parameters.AddWithValue("@SL",   m.ShiftLevel);
            cmd.Parameters.AddWithValue("@Cat",  m.CategoryName ?? "");
            cmd.Parameters.AddWithValue("@E1",   m.Event1 ?? "");
            cmd.Parameters.AddWithValue("@E2",   m.Event2 ?? "");
            cmd.Parameters.AddWithValue("@E3",   m.Event3 ?? "");
            cmd.Parameters.AddWithValue("@E4",   m.Event4 ?? "");
            cmd.Parameters.AddWithValue("@E5",   m.Event5 ?? "");
            cmd.Parameters.AddWithValue("@Desc", desc);
            cmd.Parameters.AddWithValue("@PLU",  m.PluPrefix);
        }
    }

    public class ShiftKeyDto
    {
        public bool   IsNew               { get; set; }
        public string KeyName             { get; set; } = "";
        public string OriginalKeyName     { get; set; } = "";
        public string OriginalCategoryName{ get; set; } = "";
        public int    ShiftLevel          { get; set; } = 1;
        public string CategoryName        { get; set; } = "";
        public string Event1              { get; set; } = "";
        public string Event2              { get; set; } = "";
        public string Event3              { get; set; } = "";
        public string Event4              { get; set; } = "";
        public string Event5              { get; set; } = "";
        public string Description         { get; set; } = "";
        public int    PluPrefix           { get; set; } = 0;
    }

    public class ShiftKeyDeleteDto
    {
        public string KeyName      { get; set; } = "";
        public string CategoryName { get; set; } = "";
    }
}
