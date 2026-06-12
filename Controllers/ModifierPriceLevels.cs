using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MSaleWebServer.Models;
using System.Configuration;

namespace MSaleWebServer.Controllers
{
    public class ModifierPriceLevels : Controller
    {
        private readonly IConfiguration _configuration;

        public ModifierPriceLevels(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public IActionResult GetModifierPriceLevels()
        {
            List<ModifierPriceLevelDto> list = new List<ModifierPriceLevelDto>();

            using (SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(@"
            SELECT [Modifier Title], [Prefix], [Price], [Level], [Modifier Category]
            FROM [Price Levels]
            ORDER BY [Modifier Title], [Level], [Modifier Category]", cn);

                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        list.Add(new ModifierPriceLevelDto
                        {
                            ModifierTitle = dr["Modifier Title"].ToString(),
                            Prefix = dr["Prefix"].ToString(),
                            Price = dr["Price"] == DBNull.Value ? 0 : Convert.ToDecimal(dr["Price"]),
                            Level = dr["Level"] == DBNull.Value ? 0 : Convert.ToInt32(dr["Level"]),
                            ModifierCategory = dr["Modifier Category"].ToString()
                        });
                    }
                }
            }

            return new JsonResult(list);
        }

        [HttpPost]
        public IActionResult SaveModifierPriceLevels([FromBody] List<ModifierPriceLevelDto> rows)
        {
            using (SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlTransaction tr = cn.BeginTransaction();

                try
                {
                    SqlCommand delete = new SqlCommand("DELETE FROM [Price Levels]", cn, tr);
                    delete.ExecuteNonQuery();

                    foreach (var r in rows ?? new List<ModifierPriceLevelDto>())
                    {
                        if (string.IsNullOrWhiteSpace(r.ModifierTitle) &&
                            string.IsNullOrWhiteSpace(r.ModifierCategory))
                            continue;

                        SqlCommand insert = new SqlCommand(@"
                    INSERT INTO [Price Levels]
                    (
                        [Modifier Title],
                        [Prefix],
                        [Price],
                        [Level],
                        [Modifier Category]
                    )
                    VALUES
                    (
                        @ModifierTitle,
                        @Prefix,
                        @Price,
                        @Level,
                        @ModifierCategory
                    )", cn, tr);

                        insert.Parameters.AddWithValue("@ModifierTitle", r.ModifierTitle ?? "");
                        insert.Parameters.AddWithValue("@Prefix", r.Prefix ?? "");
                        insert.Parameters.AddWithValue("@Price", r.Price);
                        insert.Parameters.AddWithValue("@Level", r.Level);
                        insert.Parameters.AddWithValue("@ModifierCategory", r.ModifierCategory ?? "");

                        insert.ExecuteNonQuery();
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Rollback();
                    throw;
                }
            }

            return new JsonResult(new { success = true, message = "Modifier price levels saved." });
        }
    }
}
