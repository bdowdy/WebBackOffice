using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    public class TableLayoutController : Controller
    {
        private readonly IConfiguration _configuration;
        public TableLayoutController(IConfiguration configuration) => _configuration = configuration;

        [HttpGet]
        public IActionResult Index() => View("~/Views/MSaleAdmin/TableLayout.cshtml");

        [HttpGet]
        public IActionResult GetRooms()
        {
            var list = new List<string>();
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();
            using var cmd = new SqlCommand(@"
                SELECT DISTINCT ISNULL([Layout Name],'') AS [Layout Name]
                FROM [Layout]
                WHERE LTRIM(RTRIM(ISNULL([Layout Name],''))) <> ''
                ORDER BY [Layout Name]", cn);
            using var dr = cmd.ExecuteReader();
            while (dr.Read()) list.Add(dr["Layout Name"].ToString() ?? "");
            return Json(list);
        }

        [HttpGet]
        public IActionResult GetRoom(string name)
        {
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            cn.Open();
            using var cmd = new SqlCommand(@"
                SELECT
                    ISNULL([Table Name],'')   AS [Table Name],
                    ISNULL([Layout Name],'')  AS [Layout Name],
                    ISNULL([Picture],'')      AS [Picture],
                    ISNULL([background],'')   AS [background],
                    ISNULL([Width],80)        AS [Width],
                    ISNULL([Height],80)       AS [Height],
                    ISNULL([Top],50)          AS [Top],
                    ISNULL([Left],50)         AS [Left],
                    ISNULL([Party Number],0)  AS [Party Number]
                FROM [Layout]
                WHERE [Layout Name] = @Name
                ORDER BY [Index Number]", cn);
            cmd.Parameters.AddWithValue("@Name", name ?? "");
            using var dr = cmd.ExecuteReader();
            var tables = new List<object>();
            string bg = "none";
            while (dr.Read())
            {
                bg = dr["background"].ToString() ?? "none";
                tables.Add(new {
                    tableName   = dr["Table Name"].ToString(),
                    picture     = dr["Picture"].ToString(),
                    width       = Convert.ToInt32(dr["Width"]),
                    height      = Convert.ToInt32(dr["Height"]),
                    top         = Convert.ToInt32(dr["Top"]),
                    left        = Convert.ToInt32(dr["Left"]),
                    partyNumber = Convert.ToInt32(dr["Party Number"])
                });
            }
            return Json(new { background = bg, tables });
        }

        [HttpPost]
        public IActionResult SaveRoom([FromBody] SaveRoomDto model)
        {
            try
            {
                if (model == null || string.IsNullOrWhiteSpace(model.LayoutName))
                    return Json(new { success = false, message = "Room name is required." });

                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();
                using var tran = cn.BeginTransaction();
                try
                {
                    using var del = new SqlCommand(
                        "DELETE FROM [Layout] WHERE [Layout Name] = @Name", cn, tran);
                    del.Parameters.AddWithValue("@Name", model.LayoutName.Trim());
                    del.ExecuteNonQuery();

                    for (int i = 0; i < model.Tables.Count; i++)
                    {
                        var t = model.Tables[i];
                        using var ins = new SqlCommand(@"
                            INSERT INTO [Layout]
                                ([Table Name],[Layout Name],[Job Code],[Index Number],
                                 [Picture],[background],[Width],[Height],[Top],[Left],
                                 [Party Number],[Assigned To])
                            VALUES
                                (@TN,@LN,@LN,@Idx,@Pic,@BG,@W,@H,@T,@L,@PN,'')", cn, tran);
                        ins.Parameters.AddWithValue("@TN",  t.TableName ?? "");
                        ins.Parameters.AddWithValue("@LN",  model.LayoutName.Trim());
                        ins.Parameters.AddWithValue("@Idx", i);
                        ins.Parameters.AddWithValue("@Pic", t.Picture ?? "");
                        ins.Parameters.AddWithValue("@BG",  model.Background ?? "none");
                        ins.Parameters.AddWithValue("@W",   t.Width);
                        ins.Parameters.AddWithValue("@H",   t.Height);
                        ins.Parameters.AddWithValue("@T",   t.Top);
                        ins.Parameters.AddWithValue("@L",   t.Left);
                        ins.Parameters.AddWithValue("@PN",  t.PartyNumber);
                        ins.ExecuteNonQuery();
                    }
                    tran.Commit();
                    return Json(new { success = true, message = "Room saved." });
                }
                catch
                {
                    tran.Rollback();
                    throw;
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public IActionResult DeleteRoom([FromBody] DeleteRoomDto model)
        {
            try
            {
                if (model == null || string.IsNullOrWhiteSpace(model.LayoutName))
                    return Json(new { success = false, message = "Room name required." });
                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();
                using var cmd = new SqlCommand(
                    "DELETE FROM [Layout] WHERE [Layout Name] = @Name", cn);
                cmd.Parameters.AddWithValue("@Name", model.LayoutName.Trim());
                cmd.ExecuteNonQuery();
                return Json(new { success = true, message = "Room deleted." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }
    }

    public class LayoutTableItemDto
    {
        public string TableName  { get; set; } = "";
        public string Picture    { get; set; } = "square4";
        public int    Width      { get; set; } = 80;
        public int    Height     { get; set; } = 80;
        public int    Top        { get; set; } = 50;
        public int    Left       { get; set; } = 50;
        public int    PartyNumber{ get; set; } = 0;
    }

    public class SaveRoomDto
    {
        public string LayoutName { get; set; } = "";
        public string Background { get; set; } = "none";
        public List<LayoutTableItemDto> Tables { get; set; } = new();
    }

    public class DeleteRoomDto
    {
        public string LayoutName { get; set; } = "";
    }
}
