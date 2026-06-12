using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MSaleWebServer.Models;

namespace MSaleWebServer.Controllers

{
    public class ModifierMaintenanceController : Controller
    {
        private readonly IConfiguration _configuration;

        public ModifierMaintenanceController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View("~/Views/MSaleAdmin/ModifierMaintenance.cshtml");
        }

        [HttpGet]
        public IActionResult ModifierMaintenance()
        {
            return View("~/Views/MSaleAdmin/ModifierMaintenance.cshtml");
        }

        [HttpGet]
        public IActionResult GetModifierNames()
        {
            List<string> list = new();

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            SqlCommand cmd = new SqlCommand(@"
                SELECT [Modifiers]
                FROM [Modifier Names]
                WHERE [Modifiers] IS NOT NULL
                  AND LTRIM(RTRIM([Modifiers])) <> ''
                ORDER BY [Modifiers]", cn);

            using SqlDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                list.Add(dr["Modifiers"].ToString() ?? "");
            }

            return Json(list);
        }

        [HttpGet]
        public IActionResult GetModifierItems(string modifierName)
        {
            List<ModifierItemDto> list = new();

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            SqlCommand cmd = new SqlCommand(@"
                SELECT [Modifier Category], [POS]
                FROM [Modifiers]
                WHERE [Modifier Name] = @ModifierName
                ORDER BY [POS], [Modifier Category]", cn);

            cmd.Parameters.AddWithValue("@ModifierName", modifierName ?? "");

            using SqlDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                list.Add(new ModifierItemDto
                {
                    ModifierCategory = dr["Modifier Category"].ToString() ?? "",
                    Pos = dr["POS"] == DBNull.Value ? 0 : Convert.ToInt32(dr["POS"])
                });
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveModifierName([FromBody] ModifierNameDto model)
        {
            if (model == null || string.IsNullOrWhiteSpace(model.ModifierName))
                return Json(new { success = false, message = "Modifier name is required." });

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            SqlCommand check = new SqlCommand(@"
                SELECT COUNT(*)
                FROM [Modifier Names]
                WHERE [Modifiers] = @Name", cn);

            check.Parameters.AddWithValue("@Name", model.ModifierName.Trim());

            int count = Convert.ToInt32(check.ExecuteScalar());

            if (count == 0)
            {
                SqlCommand insert = new SqlCommand(@"
                    INSERT INTO [Modifier Names] ([Modifiers])
                    VALUES (@Name)", cn);

                insert.Parameters.AddWithValue("@Name", model.ModifierName.Trim());
                insert.ExecuteNonQuery();
            }

            return Json(new { success = true, message = "Modifier saved." });
        }

        [HttpPost]
        public IActionResult DeleteModifierName([FromBody] ModifierNameDto model)
        {
            if (model == null || string.IsNullOrWhiteSpace(model.ModifierName))
                return Json(new { success = false, message = "Modifier name is required." });

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            using SqlTransaction tr = cn.BeginTransaction();

            try
            {
                SqlCommand deleteItems = new SqlCommand(@"
                    DELETE FROM [Modifiers]
                    WHERE [Modifier Name] = @Name", cn, tr);

                deleteItems.Parameters.AddWithValue("@Name", model.ModifierName.Trim());
                deleteItems.ExecuteNonQuery();

                SqlCommand deleteName = new SqlCommand(@"
                    DELETE FROM [Modifier Names]
                    WHERE [Modifiers] = @Name", cn, tr);

                deleteName.Parameters.AddWithValue("@Name", model.ModifierName.Trim());
                deleteName.ExecuteNonQuery();

                tr.Commit();

                return Json(new { success = true, message = "Modifier deleted." });
            }
            catch (Exception ex)
            {
                tr.Rollback();
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public IActionResult SaveModifierItems([FromBody] SaveModifierItemsDto model)
        {
            if (model == null || string.IsNullOrWhiteSpace(model.ModifierName))
                return Json(new { success = false, message = "Modifier name is required." });

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            using SqlTransaction tr = cn.BeginTransaction();

            try
            {
                SqlCommand ensureName = new SqlCommand(@"
                    IF NOT EXISTS (
                        SELECT 1 FROM [Modifier Names]
                        WHERE [Modifiers] = @Name
                    )
                    INSERT INTO [Modifier Names] ([Modifiers])
                    VALUES (@Name)", cn, tr);

                ensureName.Parameters.AddWithValue("@Name", model.ModifierName.Trim());
                ensureName.ExecuteNonQuery();

                SqlCommand deleteOld = new SqlCommand(@"
                    DELETE FROM [Modifiers]
                    WHERE [Modifier Name] = @Name", cn, tr);

                deleteOld.Parameters.AddWithValue("@Name", model.ModifierName.Trim());
                deleteOld.ExecuteNonQuery();

                int pos = 0;

                foreach (var item in model.Items ?? new List<ModifierItemDto>())
                {
                    if (string.IsNullOrWhiteSpace(item.ModifierCategory))
                        continue;

                    SqlCommand insert = new SqlCommand(@"
                        INSERT INTO [Modifiers]
                        (
                            [Modifier Name],
                            [Modifier Category],
                            [POS]
                        )
                        VALUES
                        (
                            @ModifierName,
                            @ModifierCategory,
                            @POS
                        )", cn, tr);

                    insert.Parameters.AddWithValue("@ModifierName", model.ModifierName.Trim());
                    insert.Parameters.AddWithValue("@ModifierCategory", item.ModifierCategory.Trim());
                    insert.Parameters.AddWithValue("@POS", pos);

                    insert.ExecuteNonQuery();
                    pos++;
                }

                tr.Commit();

                return Json(new { success = true, message = "Modifier items saved." });
            }
            catch (Exception ex)
            {
                tr.Rollback();
                return Json(new { success = false, message = ex.Message });
            }
        }
    }
}