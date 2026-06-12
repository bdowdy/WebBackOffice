using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/menu-categories")]
    public class MenuCategoriesController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public MenuCategoriesController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult GetMenuCategories()
        {
            try
            {
                string cs = _configuration.GetConnectionString("DefaultConnection");

                var list = new List<MenuCategoryDto>();

                using var cn = new SqlConnection(cs);
                cn.Open();

                string sql = @"
SELECT
    [Category Name],
    [Sort Level],
    [Button Level],
    [Mandatory],
    [Disable Mods],
    [Tip Exempt],
    [Prompt],
    [Meal Period],
    [BitMap],
    [Alias]
FROM [dbo].[Category Information]
ORDER BY [Sort Level], [Category Name]";

                using var cmd = new SqlCommand(sql, cn);
                using var rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    list.Add(new MenuCategoryDto
                    {
                        CategoryName = rdr["Category Name"]?.ToString() ?? "",
                        SortOrder = rdr["Sort Level"] != DBNull.Value ? Convert.ToInt32(rdr["Sort Level"]) : 0,
                        ButtonLevel = rdr["Button Level"] != DBNull.Value ? Convert.ToInt32(rdr["Button Level"]) : 0,                        
                        IsMandatory = (rdr["Mandatory"]?.ToString().Trim().ToUpper() == "YES"),
                        DisableModifiers = (rdr["Disable Mods"]?.ToString().Trim().ToUpper() == "YES"),
                        TipExempt = (rdr["Tip Exempt"]?.ToString().Trim().ToUpper() == "YES"),
                        PromptText = rdr["Prompt"]?.ToString() ?? "",
                        MealPeriod = rdr["Meal Period"]?.ToString() ?? "",
                        Image = rdr["BitMap"]?.ToString() ?? "",
                        Alias = rdr["Alias"]?.ToString() ?? ""
                    });
                }

                return Ok(list);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Error: " + ex.Message);
            }
        }
    }

    public class MenuCategoryDto
    {
        public string CategoryName { get; set; } = "";
        public int SortOrder { get; set; }
        public int ButtonLevel { get; set; }
        public bool IsMandatory { get; set; }
        public bool DisableModifiers { get; set; }
        public bool TipExempt { get; set; }
        public string PromptText { get; set; } = "";
        public string MealPeriod { get; set; } = "";
        public string Image { get; set; } = "";
        public string Alias { get; set; } = "";
    }
}