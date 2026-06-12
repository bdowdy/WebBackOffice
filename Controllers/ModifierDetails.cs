using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MSaleWebServer.Models;

namespace MSaleWebServer.Controllers
{
    public class ModifierDetailsController : Controller
    {
        private readonly IConfiguration _configuration;

        public ModifierDetailsController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public IActionResult GetModifierNames()
        {
            var list = new List<string>();

            using var cn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            using var cmd = new SqlCommand(@"
                SELECT [Modifiers]
                FROM [Modifier Names]
                WHERE [Modifiers] IS NOT NULL
                  AND LTRIM(RTRIM([Modifiers])) <> ''
                ORDER BY [Modifiers]", cn);

            using var dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                list.Add(dr["Modifiers"].ToString() ?? "");
            }

            return Json(list);
        }

        [HttpGet]
        public IActionResult GetSystemPrinters()
        {
            var list = new List<string>();

            using var cn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            using var cmd = new SqlCommand(@"
                SELECT DISTINCT [Printer Name]
                FROM [System Printers]
                WHERE [Printer Name] IS NOT NULL
                  AND LTRIM(RTRIM([Printer Name])) <> ''
                ORDER BY [Printer Name]", cn);

            using var dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                list.Add(dr["Printer Name"].ToString() ?? "");
            }

            return Json(list);
        }

        [HttpGet]
        public IActionResult GetModifierDetailItems(string modifierName)
        {
            var list = new List<ModifierDetailDto>();

            using var cn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            using var cmd = new SqlCommand(@"
                SELECT
                    [Modifier Name],
                    [Modifier Category],
                    ISNULL([Modifier Cost], 0) AS [Modifier Cost],
                    ISNULL([Sales Tax], 0) AS [Sales Tax],
                    ISNULL([Special Tax], 0) AS [Special Tax],
                    ISNULL([Liquor Tax], 0) AS [Liquor Tax],
                    ISNULL([Item Cost], 0) AS [Item Cost],
                    ISNULL([Remote Printer 1], '') AS [Remote Printer 1],
                    ISNULL([Remote Printer 2], '') AS [Remote Printer 2],
                    ISNULL([Remote Printer 3], '') AS [Remote Printer 3],
                    ISNULL([Remote Printer 4], '') AS [Remote Printer 4],
                    ISNULL([Remote Printer 5], '') AS [Remote Printer 5],
                    ISNULL([Remote Printer 6], '') AS [Remote Printer 6],
                    ISNULL([Price Level], '') AS [Price Level],
                    ISNULL([Alt Modifier], '') AS [Alt Modifier],
                    ISNULL([Selection Quantity], 0) AS [Selection Quantity],
                    ISNULL([Alias], '') AS [Alias],
                    ISNULL([ComesWith], '') AS [ComesWith],
                    ISNULL([PLU], '') AS [PLU]
                FROM [Modifiers]
                WHERE [Modifier Name] = @ModifierName
                ORDER BY ISNULL([POS], 0), [Modifier Category]", cn);

            cmd.Parameters.AddWithValue("@ModifierName", modifierName ?? "");

            using var dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                list.Add(new ModifierDetailDto
                {
                    IsNewItem = false,

                    ModifierName = dr["Modifier Name"].ToString() ?? "",
                    ModifierCategory = dr["Modifier Category"].ToString() ?? "",

                    OriginalModifierName = dr["Modifier Name"].ToString() ?? "",
                    OriginalModifierCategory = dr["Modifier Category"].ToString() ?? "",

                    ModifierCost = Convert.ToDecimal(dr["Modifier Cost"]),
                    ItemCost = Convert.ToDecimal(dr["Item Cost"]),

                    SalesTax = Convert.ToInt32(dr["Sales Tax"]),
                    SpecialTax = Convert.ToInt32(dr["Special Tax"]),
                    LiquorTax = Convert.ToInt32(dr["Liquor Tax"]),

                    RemotePrinter1 = dr["Remote Printer 1"].ToString() ?? "",
                    RemotePrinter2 = dr["Remote Printer 2"].ToString() ?? "",
                    RemotePrinter3 = dr["Remote Printer 3"].ToString() ?? "",
                    RemotePrinter4 = dr["Remote Printer 4"].ToString() ?? "",
                    RemotePrinter5 = dr["Remote Printer 5"].ToString() ?? "",
                    RemotePrinter6 = dr["Remote Printer 6"].ToString() ?? "",

                    PriceLevel = dr["Price Level"].ToString() ?? "",
                    AltModifier = dr["Alt Modifier"].ToString() ?? "",
                    SelectionQuantity = Convert.ToInt32(dr["Selection Quantity"]),
                    Alias = dr["Alias"].ToString() ?? "",
                    ComesWith = dr["ComesWith"].ToString() ?? "",
                    PLU = dr["PLU"].ToString() ?? ""
                });
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveModifierDetailItem([FromBody] ModifierDetailDto model)
        {
            try
            {
                if (model == null)
                    return Json(new { success = false, message = "No data received." });

                if (string.IsNullOrWhiteSpace(model.ModifierName))
                    return Json(new { success = false, message = "Modifier name is required." });

                if (string.IsNullOrWhiteSpace(model.ModifierCategory))
                    return Json(new { success = false, message = "Modifier category is required." });

                using var cn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                cn.Open();

                if (model.IsNewItem)
                {
                    using var check = new SqlCommand(@"
                        SELECT COUNT(*)
                        FROM [Modifiers]
                        WHERE [Modifier Name] = @ModifierName
                          AND [Modifier Category] = @ModifierCategory", cn);

                    check.Parameters.AddWithValue("@ModifierName", model.ModifierName.Trim());
                    check.Parameters.AddWithValue("@ModifierCategory", model.ModifierCategory.Trim());

                    int exists = Convert.ToInt32(check.ExecuteScalar());

                    if (exists > 0)
                        return Json(new { success = false, message = "This modifier item already exists." });

                    using var insert = new SqlCommand(@"
                        INSERT INTO [Modifiers]
                        (
                            [Modifier Name],
                            [Modifier Category],
                            [Modifier Cost],
                            [Sales Tax],
                            [Special Tax],
                            [Liquor Tax],
                            [Item Cost],
                            [Remote Printer 1],
                            [Remote Printer 2],
                            [Remote Printer 3],
                            [Remote Printer 4],
                            [Remote Printer 5],
                            [Remote Printer 6],
                            [Price Level],
                            [Alt Modifier],
                            [Selection Quantity],
                            [Alias],
                            [ComesWith],
                            [PLU],
                            [POS]
                        )
                        VALUES
                        (
                            @ModifierName,
                            @ModifierCategory,
                            @ModifierCost,
                            @SalesTax,
                            @SpecialTax,
                            @LiquorTax,
                            @ItemCost,
                            @RemotePrinter1,
                            @RemotePrinter2,
                            @RemotePrinter3,
                            @RemotePrinter4,
                            @RemotePrinter5,
                            @RemotePrinter6,
                            @PriceLevel,
                            @AltModifier,
                            @SelectionQuantity,
                            @Alias,
                            @ComesWith,
                            @PLU,
                            ISNULL((SELECT MAX([POS]) + 1 FROM [Modifiers] WHERE [Modifier Name] = @ModifierName), 0)
                        )", cn);

                    AddParameters(insert, model);

                    int rows = insert.ExecuteNonQuery();

                    return Json(new
                    {
                        success = rows == 1,
                        message = rows == 1 ? "New modifier item added." : "Insert failed."
                    });
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(model.OriginalModifierName) ||
                        string.IsNullOrWhiteSpace(model.OriginalModifierCategory))
                    {
                        return Json(new
                        {
                            success = false,
                            message = "Original modifier information is missing. Reload the item and try again."
                        });
                    }

                    using var update = new SqlCommand(@"
                        UPDATE [Modifiers]
                        SET
                            [Modifier Name] = @ModifierName,
                            [Modifier Category] = @ModifierCategory,
                            [Modifier Cost] = @ModifierCost,
                            [Sales Tax] = @SalesTax,
                            [Special Tax] = @SpecialTax,
                            [Liquor Tax] = @LiquorTax,
                            [Item Cost] = @ItemCost,
                            [Remote Printer 1] = @RemotePrinter1,
                            [Remote Printer 2] = @RemotePrinter2,
                            [Remote Printer 3] = @RemotePrinter3,
                            [Remote Printer 4] = @RemotePrinter4,
                            [Remote Printer 5] = @RemotePrinter5,
                            [Remote Printer 6] = @RemotePrinter6,
                            [Price Level] = @PriceLevel,
                            [Alt Modifier] = @AltModifier,
                            [Selection Quantity] = @SelectionQuantity,
                            [Alias] = @Alias,
                            [ComesWith] = @ComesWith,
                            [PLU] = @PLU
                        WHERE [Modifier Name] = @OriginalModifierName
                          AND [Modifier Category] = @OriginalModifierCategory", cn);

                    AddParameters(update, model);

                    update.Parameters.AddWithValue("@OriginalModifierName", model.OriginalModifierName.Trim());
                    update.Parameters.AddWithValue("@OriginalModifierCategory", model.OriginalModifierCategory.Trim());

                    int rows = update.ExecuteNonQuery();

                    return Json(new
                    {
                        success = rows == 1,
                        message = rows == 1 ? "Modifier item saved." : "Update failed. Item was not found."
                    });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        private static void AddParameters(SqlCommand cmd, ModifierDetailDto model)
        {
            cmd.Parameters.AddWithValue("@ModifierName", model.ModifierName?.Trim() ?? "");
            cmd.Parameters.AddWithValue("@ModifierCategory", model.ModifierCategory?.Trim() ?? "");
            cmd.Parameters.AddWithValue("@ModifierCost", model.ModifierCost);
            cmd.Parameters.AddWithValue("@SalesTax", model.SalesTax);
            cmd.Parameters.AddWithValue("@SpecialTax", model.SpecialTax);
            cmd.Parameters.AddWithValue("@LiquorTax", model.LiquorTax);
            cmd.Parameters.AddWithValue("@ItemCost", model.ItemCost);

            cmd.Parameters.AddWithValue("@RemotePrinter1", model.RemotePrinter1 ?? "");
            cmd.Parameters.AddWithValue("@RemotePrinter2", model.RemotePrinter2 ?? "");
            cmd.Parameters.AddWithValue("@RemotePrinter3", model.RemotePrinter3 ?? "");
            cmd.Parameters.AddWithValue("@RemotePrinter4", model.RemotePrinter4 ?? "");
            cmd.Parameters.AddWithValue("@RemotePrinter5", model.RemotePrinter5 ?? "");
            cmd.Parameters.AddWithValue("@RemotePrinter6", model.RemotePrinter6 ?? "");

            cmd.Parameters.AddWithValue("@PriceLevel", model.PriceLevel ?? "");
            cmd.Parameters.AddWithValue("@AltModifier", model.AltModifier ?? "");
            cmd.Parameters.AddWithValue("@SelectionQuantity", model.SelectionQuantity);
            cmd.Parameters.AddWithValue("@Alias", model.Alias ?? "");
            cmd.Parameters.AddWithValue("@ComesWith", model.ComesWith ?? "");
            cmd.Parameters.AddWithValue("@PLU", model.PLU ?? "");
        }
    }
}