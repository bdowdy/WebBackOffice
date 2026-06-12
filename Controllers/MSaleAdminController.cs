using Dapper;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using MSaleWebServer.Models;
using System.Configuration;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace MSaleWebServer.Controllers
{
    public partial class MSaleAdminController : Controller
    {
        private readonly IConfiguration _configuration;

        public MSaleAdminController(IConfiguration configuration)
        {
            _configuration = configuration;
        }


        public IActionResult Index()
        {
            return View();
        }
        public IActionResult TenderAdmin()
        {
            return View("~/Views/MSaleAdmin/TenderAdmin.cshtml");
        }

        public IActionResult CouponAssignments()
        {
            return View("~/Views/MSaleAdmin/CouponAssignments.cshtml");
        }

        
        public IActionResult Edit(string id)
        {
            var model = new DiscountEditModel();
            ViewBag.ActiveDays = GetDiscountDays(model.DiscountName);
            ViewBag.SalesTaxOptions = GetSalesTaxOptions();
            ViewBag.SpecialTaxOptions = GetSpecialTaxOptions();
            ViewBag.DiscountCategories = GetDiscountCategories();
            ViewBag.DiscountTypes = GetDiscountTypes();
            string cs = _configuration.GetConnectionString("DefaultConnection");

            using var cn = new SqlConnection(cs);
            cn.Open();

            string sql = @"
SELECT
 [Discount Name],
 [Discount Type],
 [Discount Amount],
 [Access Level],
 [Start Time],
 [End Time],
 [Expires],
[Allow Only One],
[After Print],
[Auto Match],
[Auto Deduct],
[Tracking],
[Search Number]
FROM Discounts
WHERE [Discount Name]=@name";

            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@name", id ?? "");

            using var rdr = cmd.ExecuteReader();

            if (rdr.Read())
            {
                model.DiscountName = rdr["Discount Name"]?.ToString() ?? "";
                model.DiscountType = rdr["Discount Type"]?.ToString() ?? "";
                model.DiscountAmount = rdr["Discount Amount"] == DBNull.Value ? 0 : Convert.ToDecimal(rdr["Discount Amount"]);
                model.AccessLevel = rdr["Access Level"]?.ToString() ?? "";
                model.StartTime = rdr["Start Time"]?.ToString() ?? "";

                model.EndTime = rdr["End Time"]?.ToString() ?? "";
                model.Expires = rdr["Expires"]?.ToString() ?? "";
                model.AllowOnlyOne = ToBool(rdr["Allow Only One"]);
                model.AfterPrint = ToBool(rdr["After Print"]);
                model.AutoMatch = ToBool(rdr["Auto Match"]);
                model.AutoDeduct = ToBool(rdr["Auto Deduct"]);
                model.Tracking = ToBool(rdr["Tracking"]);

                model.EventOnly = ToBool(rdr["Search Number"]);
            }

            return View(model);
        }
        public IActionResult Create()
        {

            ViewBag.SalesTaxOptions = GetSalesTaxOptions();
            ViewBag.SpecialTaxOptions = GetSpecialTaxOptions();
            ViewBag.DiscountCategories = GetDiscountCategories();
            ViewBag.DiscountTypes = GetDiscountTypes();
            ViewBag.ActiveDays = GetDiscountDays("");
            return View("Edit", new DiscountEditModel());
        }
        public IActionResult DailySpecials()
        {
            return View();
        }

        [HttpGet]
        public IActionResult GetDailySpecial(string day)
        {
            try
            {
                using var conn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                var row = conn.QueryFirstOrDefault<DailySpecialDto>(@"
            SELECT
                [Day] AS Day,
                [Special 1] AS Special1,
                [Special 2] AS Special2,
                [Special 3] AS Special3,
                [Special 4] AS Special4,
                [Flash Mess 1] AS FlashMess1,
                [Flash Mess 2] AS FlashMess2,
                [Flash Mess 3] AS FlashMess3,
                [Flash Mess 4] AS FlashMess4
            FROM [Daily Specials]
            WHERE [Day] = @Day",
                    new { Day = day });

                if (row == null)
                {
                    row = new DailySpecialDto
                    {
                        Day = day
                    };
                }

                return Json(row);
            }
            catch (Exception ex)
            {
                return Json(new { error = ex.Message });
            }
        }
        [HttpPost]
        public IActionResult SaveDailySpecial([FromBody] DailySpecialDto data)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            conn.Execute("DELETE FROM [Daily Specials] WHERE [Day] = @Day",
                new { data.Day });

            conn.Execute(@"
        INSERT INTO [Daily Specials]
        ([Special 1], [Special 2], [Special 3], [Special 4],
         [Flash Mess 1], [Flash Mess 2], [Flash Mess 3], [Flash Mess 4], [Day])
        VALUES
        (@Special1, @Special2, @Special3, @Special4,
         @FlashMess1, @FlashMess2, @FlashMess3, @FlashMess4, @Day)",
                data);

            return Json(new { message = "Saved" });
        }
        [HttpPost]
        public IActionResult DeleteDailySpecial(string day)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            conn.Execute("DELETE FROM [Daily Specials] WHERE [Day] = @day",
                new { day });

            return Json(new { message = "Removed" });
        }

        [HttpGet]
        public IActionResult GetOrderButtons()
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var data = conn.Query(@"
        SELECT
            [Idx No] AS idxNo,
            [Caption Name] AS captionName,
            [Spare Left] AS spareLeft,
            [Spare Top] AS spareTop
        FROM [Order Buttons]
        ORDER BY [Idx No]
    ").ToList();

            return Json(data);
        }


        [HttpPost]
        public IActionResult SaveOrderButtons([FromBody] List<OrderButtonDto> rows)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            foreach (var r in rows)
            {
                conn.Execute(@"
            UPDATE [Order Buttons]
            SET
                [Caption Name] = @CaptionName,
                [Spare Left] = @SpareLeft,
                [Spare Top] = @SpareTop
            WHERE [Idx No] = @IdxNo", r);
            }

            return Json(new { message = "Saved" });
        }

        public IActionResult OrderButtons()
        {
            return View();
        }
        public IActionResult MenuAccess()
        {
            return View();
        }
        public IActionResult MealPeriods()
        {
            return View("~/Views/MSaleAdmin/MealPeriods.cshtml");
        }

        [HttpGet]
        public IActionResult GetTimedEvents()
        {
            List<string> list = new List<string>();

            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"SELECT [Event Name]
              FROM [Timed Pricing]
              ORDER BY [Event Name]", cn);

                SqlDataReader rd = cmd.ExecuteReader();

                while (rd.Read())
                {
                    list.Add(rd["Event Name"].ToString());
                }
            }

            return Json(list);
        }
        [HttpGet]
        public IActionResult GetCategories()
        {
            List<string> list = new List<string>();

            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"SELECT [Category Name]
              FROM [Category Information]
              ORDER BY [Category Name]", cn);

                SqlDataReader rd = cmd.ExecuteReader();

                while (rd.Read())
                {
                    list.Add(rd["Category Name"].ToString());
                }
            }

            return Json(list);
        }
        [HttpPost]
        public IActionResult SaveMealPeriod(
    [FromBody] MealPeriodModel model)
        {
            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand del = new SqlCommand(
                    @"DELETE FROM [Meal Period]
              WHERE [Period Name] = @name", cn);

                del.Parameters.AddWithValue("@name",
                    model.PeriodName);

                del.ExecuteNonQuery();

                SqlCommand ins = new SqlCommand(
                    @"INSERT INTO [Meal Period]
              ([Period Name],[Period Time])
              VALUES (@name,@time)", cn);

                ins.Parameters.AddWithValue("@name",
                    model.PeriodName);

                ins.Parameters.AddWithValue("@time",
                    model.PeriodTime);

                ins.ExecuteNonQuery();
            }

            return Json(new { success = true });
        }
        [HttpPost]
        public IActionResult DeleteMealPeriod(string periodName)
        {
            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"DELETE FROM [Meal Period]
              WHERE [Period Name] = @name", cn);

                cmd.Parameters.AddWithValue("@name",
                    periodName);

                cmd.ExecuteNonQuery();
            }

            return Json(new { success = true });
        }
        [HttpGet]
        public IActionResult GetCategoryMealPeriod(
    string categoryName)
        {
            string mealPeriod = "";

            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"SELECT [Meal Period]
              FROM [Category Information]
              WHERE [Category Name] = @name", cn);

                cmd.Parameters.AddWithValue("@name",
                    categoryName);

                object result = cmd.ExecuteScalar();

                if (result != null)
                    mealPeriod = result.ToString();
            }

            return Json(new
            {
                mealPeriod = mealPeriod
            });
        }
        [HttpPost]
        public IActionResult SaveCategoryMealPeriod(
    [FromBody] CategoryMealPeriodModel model)
        {
            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"UPDATE [Category Information]
              SET [Meal Period] = @period
              WHERE [Category Name] = @name", cn);

                cmd.Parameters.AddWithValue("@period",
                    model.MealPeriod);

                cmd.Parameters.AddWithValue("@name",
                    model.CategoryName);

                cmd.ExecuteNonQuery();
            }

            return Json(new { success = true });
        }
        [HttpGet]
        public IActionResult GetMenuAccess()
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var rows = conn.Query(@"
        SELECT
            [Menu Name] AS menuName,
            [Access Required] AS accessRequired
        FROM [Menu Access]
        ORDER BY [Menu Name] ASC
    ").ToList();

            return Json(rows);
        }

        [HttpGet]
        public IActionResult GetMenuAccessItem(string menuName)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var row = conn.QueryFirstOrDefault(@"
        SELECT
            [Menu Name] AS menuName,
            [Access Required] AS accessRequired
        FROM [Menu Access]
        WHERE [Menu Name] = @menuName
    ", new { menuName });

            return Json(row ?? new { menuName, accessRequired = 0 });
        }

        [HttpPost]
        public IActionResult SaveMenuAccess([FromBody] MenuAccessDto item)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var exists = conn.ExecuteScalar<int>(@"
        SELECT COUNT(*)
        FROM [Menu Access]
        WHERE [Menu Name] = @MenuName
    ", item);

            if (exists > 0)
            {
                conn.Execute(@"
            UPDATE [Menu Access]
            SET [Access Required] = @AccessRequired
            WHERE [Menu Name] = @MenuName
        ", item);
            }
            else
            {
                conn.Execute(@"
            INSERT INTO [Menu Access]
            ([Menu Name], [Access Required])
            VALUES (@MenuName, @AccessRequired)
        ", item);
            }

            return Json(new { message = "Updated" });
        }
        public IActionResult SystemLists()
        {
            return View();
        }
        [HttpGet]
        public IActionResult GetMasterDepartments()
        {
            List<string> list = new List<string>();

            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"SELECT [Master Name]
              FROM [Master Departments]
              ORDER BY [Master Name]", cn);

                SqlDataReader rd = cmd.ExecuteReader();

                while (rd.Read())
                {
                    list.Add(rd["Master Name"].ToString());
                }
            }

            return Json(list);
        }
        [HttpGet]
        public IActionResult GetNonSaleRevenue()
        {
            List<string> list = new List<string>();

            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"SELECT [Name]
              FROM [NonSaleRevenue]
              ORDER BY [Name]", cn);

                SqlDataReader rd = cmd.ExecuteReader();

                while (rd.Read())
                {
                    list.Add(rd["Name"].ToString());
                }
            }

            return Json(list);
        }
        public class AddRevenueModel
        {
            public string Name { get; set; }
        }
        [HttpPost]
        public IActionResult AddRevenueItem(
      [FromBody] AddRevenueModel model)
        {
            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"INSERT INTO [NonSaleRevenue]
              ([Name])
              VALUES (@name)", cn);

                cmd.Parameters.AddWithValue("@name",
                    model.Name);

                cmd.ExecuteNonQuery();
            }

            return Json(new { success = true });
        }
        [HttpPost]
        public IActionResult AddMasterDepartment(
    [FromBody] AddRevenueModel model)
        {
            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"INSERT INTO [Master Departments]
              ([Master Name])
              VALUES (@name)", cn);

                cmd.Parameters.AddWithValue("@name", model.Name);

                cmd.ExecuteNonQuery();
            }

            return Json(new { success = true });
        }
        [HttpGet]
        public IActionResult GetRevenueAssignments(string revenueName)
        {
            List<string> list = new List<string>();

            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"SELECT [Link To Name]
              FROM [Revenue Assignments]
              WHERE [Revenue Name] = @name", cn);

                cmd.Parameters.AddWithValue("@name", revenueName);

                SqlDataReader rd = cmd.ExecuteReader();

                while (rd.Read())
                {
                    list.Add(rd["Link To Name"].ToString());
                }
            }

            return Json(list);
        }
        [HttpPost]
        public IActionResult SaveRevenueAssignments(
    [FromBody] RevenueAssignmentModel model)
        {
            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand del = new SqlCommand(
                    @"DELETE FROM [Revenue Assignments]
              WHERE [Revenue Name] = @name", cn);

                del.Parameters.AddWithValue("@name",
                    model.RevenueName);

                del.ExecuteNonQuery();

                foreach (var item in model.LinkToNames)
                {
                    SqlCommand ins = new SqlCommand(
                        @"INSERT INTO [Revenue Assignments]
                  ([Revenue Name],[Link To Name],[Link Type])
                  VALUES (@r,@l,@t)", cn);

                    ins.Parameters.AddWithValue("@r",
                        model.RevenueName);

                    ins.Parameters.AddWithValue("@l",
                        item);

                    ins.Parameters.AddWithValue("@t",
                        model.LinkType);

                    ins.ExecuteNonQuery();
                }
            }

            return Json(new { success = true });
        }
        public IActionResult RevenueAssignments()
        {
            return View();
        }
        public IActionResult AccessLevels()
        {
            return View();
        }
        [HttpGet]
        public IActionResult GetAccessLevelNames()
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var rows = conn.Query(@"
        SELECT distinct [Access Name] AS accessName            
        FROM [Access Levels]
        ORDER BY [Access Name]
    ").ToList();

            return Json(rows);
        }

        [HttpGet]
        public IActionResult GetAccessLevel(string accessName)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var row = conn.QueryFirstOrDefault(@"
        SELECT
            [Access Name] AS accessName,
            [Access Level] AS accessLevel
        FROM [Access Levels]
        WHERE [Access Name] = @accessName
    ", new { accessName });

            return Json(row ?? new { accessName, accessLevel = 1 });
        }

        [HttpPost]
        public IActionResult SaveAccessLevel([FromBody] AccessLevelDto item)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var exists = conn.ExecuteScalar<int>(@"
        SELECT COUNT(*)
        FROM [Access Levels]
        WHERE [Access Name] = @AccessName
    ", item);

            if (exists > 0)
            {
                conn.Execute(@"
            UPDATE [Access Levels]
            SET [Access Level] = @AccessLevel
            WHERE [Access Name] = @AccessName
        ", item);
            }
            else
            {
                conn.Execute(@"
            INSERT INTO [Access Levels]
            ([Access Name], [Access Level])
            VALUES (@AccessName, @AccessLevel)
        ", item);
            }

            return Json(new { message = "Saved" });
        }

        [HttpPost]
        public IActionResult DeleteAccessLevel(string accessName)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            conn.Execute(@"
        DELETE FROM [Access Levels]
        WHERE [Access Name] = @accessName
    ", new { accessName });

            return Json(new { message = "Deleted" });
        }

        [HttpGet]
        public IActionResult GetTimedPricing()
        {
            try
            {
                using var conn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                var rows = conn.Query<TimedPricingDto>(@"
            SELECT
                [Event Name] AS EventName,
                CONVERT(varchar(5), [Start Time], 108) AS StartTime,
                CONVERT(varchar(5), [End Time], 108) AS EndTime,

                ISNULL([Monday], 0) AS Monday,
                ISNULL([Tuesday], 0) AS Tuesday,
                ISNULL([Wednesday], 0) AS Wednesday,
                ISNULL([Thursday], 0) AS Thursday,
                ISNULL([Friday], 0) AS Friday,
                ISNULL([Saturday], 0) AS Saturday,
                ISNULL([Sunday], 0) AS Sunday,

                CONVERT(varchar(10), [Start Date], 23) AS StartDate,
                CONVERT(varchar(10), [End Date], 23) AS EndDate,

                ISNULL([Price Off], 0) AS PriceOff
            FROM [Timed Pricing]
            ORDER BY [Event Name]
        ").ToList();

                return Json(rows);
            }
            catch (Exception ex)
            {
                return Json(new { error = ex.Message });
            }
        }
        public IActionResult TimedPricing()
        {
            return View();
        }
        [HttpPost]
        public IActionResult SaveTimedPricing([FromBody] List<TimedPricingDto> rows)
        {
            try
            {
                using var conn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                conn.Open();

                using var tran = conn.BeginTransaction();

                try
                {
                    conn.Execute("DELETE FROM [Timed Pricing]", transaction: tran);

                    foreach (var r in rows)
                    {
                        conn.Execute(@"
                    INSERT INTO [Timed Pricing]
                    ([Event Name],[Start Time],[End Time],
                     [Monday],[Tuesday],[Wednesday],[Thursday],[Friday],[Saturday],[Sunday],
                     [Start Date],[End Date],[Price Off])
                    VALUES
                    (@EventName,@StartTime,@EndTime,
                     @Monday,@Tuesday,@Wednesday,@Thursday,@Friday,@Saturday,@Sunday,
                     @StartDate,@EndDate,@PriceOff)",
                            new
                            {
                                EventName = r.EventName ?? "",
                                StartTime = r.StartTime ?? "",
                                EndTime = r.EndTime ?? "",

                                Monday = r.Monday ? 1 : 0,
                                Tuesday = r.Tuesday ? 1 : 0,
                                Wednesday = r.Wednesday ? 1 : 0,
                                Thursday = r.Thursday ? 1 : 0,
                                Friday = r.Friday ? 1 : 0,
                                Saturday = r.Saturday ? 1 : 0,
                                Sunday = r.Sunday ? 1 : 0,

                                StartDate = string.IsNullOrWhiteSpace(r.StartDate) ? null : r.StartDate,
                                EndDate = string.IsNullOrWhiteSpace(r.EndDate) ? null : r.EndDate,

                                PriceOff = r.PriceOff
                            },
                            transaction: tran);
                    }

                    tran.Commit();

                    return Json(new { message = "Saved " + rows.Count + " timed pricing rows." });
                }
                catch
                {
                    tran.Rollback();
                    throw;
                }
            }
            catch (Exception ex)
            {
                return Json(new { message = "ERROR: " + ex.Message });
            }
        }
        public IActionResult SubCategories()
        {
            return View();
        }
        [HttpGet]
        public IActionResult GetSubCategoryLists()
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var subCategories = conn.Query<string>(@"
        SELECT [Sub Name]
        FROM [Sub Categories]
        ORDER BY [Sub Name]
    ").ToList();

            var categories = conn.Query<string>(@"
        SELECT [Category Name]
        FROM [Category Information]
        ORDER BY [Category Name]
    ").ToList();

            var timedEvents = conn.Query<string>(@"
        SELECT [Event Name]
        FROM [Timed Pricing]
        ORDER BY [Event Name]
    ").ToList();

            return Json(new { subCategories, categories, timedEvents });
        }

        [HttpGet]
        public IActionResult GetSubCategory(string subName)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var row = conn.QueryFirstOrDefault(@"
        SELECT
            [Sub Name] AS subName,
            [Category Name] AS categoryName,
            ISNULL([Event 1], 'None') AS timedEvent1,
            ISNULL([Event 2], 'None') AS timedEvent2,
            ISNULL([Event 3], 'None') AS timedEvent3,
            ISNULL([Event 4], 'None') AS timedEvent4,
            ISNULL([Event 5], 'None') AS timedEvent5
        FROM [Sub Categories]
        WHERE [Sub Name] = @subName
    ", new { subName });

            return Json(row ?? new { subName });
        }

        [HttpPost]
        public IActionResult SaveSubCategory([FromBody] SubCategoryDto item)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var categoryExists = conn.ExecuteScalar<int>(@"
        SELECT COUNT(*)
        FROM [Category Information]
        WHERE [Category Name] = @SubName
    ", item);

            if (categoryExists > 0)
                return Json(new { message = "Cannot use same name as category." });

            var menuItemExists = conn.ExecuteScalar<int>(@"
        SELECT COUNT(*)
        FROM [Menu Items]
        WHERE [Menu Item Name] = @SubName
    ", item);

            if (menuItemExists > 0)
                return Json(new { message = "Subcategory name cannot be identical to a menu item name." });

            var exists = conn.ExecuteScalar<int>(@"
        SELECT COUNT(*)
        FROM [Sub Categories]
        WHERE [Sub Name] = @SubName
    ", item);

            if (exists > 0)
            {
                conn.Execute(@"
            UPDATE [Sub Categories]
            SET
                [Category Name] = @CategoryName,
                [Event 1] = @TimedEvent1,
                [Event 2] = @TimedEvent2,
                [Event 3] = @TimedEvent3,
                [Event 4] = @TimedEvent4,
                [Event 5] = @TimedEvent5
            WHERE [Sub Name] = @SubName
        ", item);
            }
            else
            {
                conn.Execute(@"
            INSERT INTO [Sub Categories]
            ([Sub Name], [Category Name],
             [Event 1], [Event 2], [Event 3],
             [Event 4], [Event 5])
            VALUES
            (@SubName, @CategoryName,
             @TimedEvent1, @TimedEvent2, @TimedEvent3,
             @TimedEvent4, @TimedEvent5)
        ", item);
            }

            return Json(new { message = "Saved" });
        }

        [HttpPost]
        public IActionResult DeleteSubCategory(string subName)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var usedByMenuItems = conn.ExecuteScalar<int>(@"
        SELECT COUNT(*)
        FROM [Menu Items]
        WHERE [Department Name] = @subName
    ", new { subName });

            if (usedByMenuItems > 0)
                return Json(new { message = "Menu Items are currently using this name. Cannot delete." });

            conn.Execute(@"
        DELETE FROM [Sub Categories]
        WHERE [Sub Name] = @subName
    ", new { subName });

            return Json(new { message = "Removed" });
        }

        [HttpPost]
        public IActionResult RenameSubCategory([FromBody] RenameSubCategoryDto item)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var categoryExists = conn.ExecuteScalar<int>(@"
        SELECT COUNT(*)
        FROM [Category Information]
        WHERE [Category Name] = @NewName
    ", item);

            if (categoryExists > 0)
                return Json(new { message = "Cannot use same name as category." });

            conn.Execute(@"
        UPDATE [Sub Categories]
        SET [Sub Name] = @NewName
        WHERE [Sub Name] = @OldName
    ", item);

            conn.Execute(@"
        UPDATE [Button Layout]
        SET [Menu Item] = @NewName
        WHERE [Menu Item] = @OldName
    ", item);

            conn.Execute(@"
        UPDATE [Button Layout]
        SET [Category] = @NewName
        WHERE [Category] = @OldName
    ", item);

            conn.Execute(@"
        UPDATE [Menu Items]
        SET [Department Name] = @NewName
        WHERE [Department Name] = @OldName
    ", item);

            return Json(new { message = "Update complete." });
        }
        public IActionResult ReceiptMessages()
        {
            return View();
        }
        [HttpGet]
        public IActionResult GetTicketInfo(string day)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var row = conn.QueryFirstOrDefault(@"
        SELECT
            [Line 1] AS line1,
            [Line 2] AS line2,
            [Line 3] AS line3,
            [Line 4] AS line4,
            [Day] AS day
        FROM [Ticket Info]
        WHERE [Day] = @day
    ", new { day });

            return Json(row ?? new { day, line1 = "", line2 = "", line3 = "", line4 = "" });
        }

        [HttpPost]
        public IActionResult SaveTicketInfo([FromBody] ReceiptMessageDto item)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var exists = conn.ExecuteScalar<int>(@"
        SELECT COUNT(*)
        FROM [Ticket Info]
        WHERE [Day] = @Day
    ", item);

            if (exists > 0)
            {
                conn.Execute(@"
            UPDATE [Ticket Info]
            SET
                [Line 1] = @Line1,
                [Line 2] = @Line2,
                [Line 3] = @Line3,
                [Line 4] = @Line4
            WHERE [Day] = @Day
        ", item);
            }
            else
            {
                conn.Execute(@"
            INSERT INTO [Ticket Info]
            ([Line 1], [Line 2], [Line 3], [Line 4], [Day])
            VALUES
            (@Line1, @Line2, @Line3, @Line4, @Day)
        ", item);
            }

            return Json(new { message = "Receipt message saved" });
        }
        [HttpGet]
        public IActionResult GetCreditCardMessage(string day)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var row = conn.QueryFirstOrDefault(@"
        SELECT
            [Line 1] AS line1,
            [Line 2] AS line2,
            [Line 3] AS line3,
            [Line 4] AS line4,
            [Day] AS day
        FROM [CreditCardMessage]
        WHERE [Day] = @day
    ", new { day });

            return Json(row ?? new { day, line1 = "", line2 = "", line3 = "", line4 = "" });
        }

        [HttpPost]
        public IActionResult SaveCreditCardMessage([FromBody] ReceiptMessageDto item)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var exists = conn.ExecuteScalar<int>(@"
        SELECT COUNT(*)
        FROM [CreditCardMessage]
        WHERE [Day] = @Day
    ", item);

            if (exists > 0)
            {
                conn.Execute(@"
            UPDATE [CreditCardMessage]
            SET
                [Line 1] = @Line1,
                [Line 2] = @Line2,
                [Line 3] = @Line3,
                [Line 4] = @Line4
            WHERE [Day] = @Day
        ", item);
            }
            else
            {
                conn.Execute(@"
            INSERT INTO [CreditCardMessage]
            ([Line 1], [Line 2], [Line 3], [Line 4], [Day])
            VALUES
            (@Line1, @Line2, @Line3, @Line4, @Day)
        ", item);
            }

            return Json(new { message = "Credit card slip message saved" });
        }

        [HttpPost]
        public IActionResult DeleteCreditCardMessage(string day)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            conn.Execute("DELETE FROM [CreditCardMessage] WHERE [Day] = @day", new { day });

            return Json(new { message = "Credit card slip message deleted" });
        }

        [HttpPost]
        public IActionResult DeleteTicketInfo(string day)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            conn.Execute("DELETE FROM [Ticket Info] WHERE [Day] = @day", new { day });

            return Json(new { message = "Receipt message deleted" });
        }
        public IActionResult ButtonLayout()
        {
            return View();
        }
        [HttpGet]
        public IActionResult GetButtonLayoutCategories()
        {
            List<string> list = new List<string>();

            using (SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(@"
            SELECT [Category Name]
            FROM [Category Information]
            WHERE [Category Name] IS NOT NULL
              AND LTRIM(RTRIM([Category Name])) <> ''
            ORDER BY [Category Name]", cn);

                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        list.Add(dr["Category Name"].ToString());
                    }
                }
            }

            return new JsonResult(list);
        }

        [HttpGet]
        public IActionResult GetButtonLayoutMenuItems()
        {
            List<string> list = new List<string>();

            using (SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(@"
            SELECT [Menu Item Name]
            FROM [Menu Items]
            WHERE [Menu Item Name] IS NOT NULL
              AND LTRIM(RTRIM([Menu Item Name])) <> ''
            ORDER BY [Menu Item Name]", cn);

                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        list.Add(dr["Menu Item Name"].ToString());
                    }
                }
            }

            return new JsonResult(list);
        }
        [HttpGet]
        public IActionResult GetButtonLayout(string category)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            var buttons = conn.Query<ButtonLayoutDto>(@"
        SELECT
            [Menu Item] AS MenuItem,
            [Category] AS Category,
            ISNULL([ForeCol], 16777215) AS ForeCol,
            ISNULL([BackCol], 4210752) AS BackCol,
            ISNULL([Icon], '') AS Icon,
            ISNULL([Width], 120) AS Width,
            ISNULL([Height], 55) AS Height,
            ISNULL([FontSize], 12) AS FontSize,
            ISNULL([Left], 20) AS [Left],
            ISNULL([Top], 20) AS [Top],
            ISNULL([Button Assigned], 0) AS ButtonAssigned
        FROM [Button Layout]
        WHERE [Category] = @category
        ORDER BY [Button Assigned]
    ", new { category }).ToList();

            var assigned = buttons.Select(x => x.MenuItem).ToList();

            var menuItems = conn.Query<string>(@"
        SELECT [Menu Item Name]
        FROM [Menu Items]
        WHERE ISNULL([Quick PICK], 0) = 0
          AND [Department Name] = @category
        ORDER BY [Menu Item Name]
    ", new { category }).ToList();

            var subCategories = conn.Query<string>(@"
        SELECT [Sub Name]
        FROM [Sub Categories]
        WHERE [Category Name] = @category
        ORDER BY [Sub Name]
    ", new { category }).ToList();

            var shiftKeys = conn.Query<string>(@"
        SELECT [Key Name]
        FROM [Shift Keys]
        WHERE [Category Name] = @category
        ORDER BY [Key Name]
    ", new { category }).ToList();

            var availableItems = menuItems
                .Concat(subCategories)
                .Concat(shiftKeys)
                .Where(x => !assigned.Contains(x))
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            return Json(new
            {
                buttons,
                availableItems
            });
        }
        [HttpPost]
        public IActionResult SaveButtonLayout([FromBody] ButtonLayoutSaveRequest request)
        {
            if (request == null)
                return Json(new { message = "ERROR: No data received." });

            if (string.IsNullOrWhiteSpace(request.Category))
                return Json(new { message = "ERROR: No category received." });

            if (request.Buttons == null)
                return Json(new { message = "ERROR: No buttons received." });

            try
            {
                using var conn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                conn.Open();

                using var tran = conn.BeginTransaction();

                try
                {
                    conn.Execute(@"
                DELETE FROM [Button Layout]
                WHERE [Category] = @Category
            ", new { request.Category }, tran);

                    int pos = 0;

                    foreach (var b in request.Buttons)
                    {
                        if (string.IsNullOrWhiteSpace(b.MenuItem))
                            continue;

                        conn.Execute(@"
        INSERT INTO [Button Layout]
        ([Menu Item],
         [Category],
         [ForeCol],
         [BackCol],
         [Icon],
         [Width],
         [Height],
         [FontSize],
         [Left],
         [Top],
         [Button Assigned],
         [Sound])
        VALUES
        (@MenuItem,
         @Category,
         @ForeCol,
         @BackCol,
         @Icon,
         @Width,
         @Height,
         @FontSize,
         @Left,
         @Top,
         @ButtonAssigned,
         '')",
                        new
                        {
                            MenuItem = b.MenuItem ?? "",
                            Category = request.Category,

                            // 🔥 IMPORTANT DEFAULTS
                            ForeCol = string.IsNullOrWhiteSpace(b.ForeCol) ? "16777215" : b.ForeCol,
                            BackCol = string.IsNullOrWhiteSpace(b.BackCol) ? "4210752" : b.BackCol,
                            Icon = b.Icon ?? "",

                            Width = b.Width ?? 120,
                            Height = b.Height ?? 55,
                            FontSize = b.FontSize ?? 12,
                            Left = b.Left ?? 20,
                            Top = b.Top ?? 20,

                            ButtonAssigned = pos
                        },
                        tran);

                        pos++;
                    }
                    tran.Commit();

                    return Json(new { message = "Saved " + pos + " buttons." });
                }
                catch
                {
                    tran.Rollback();
                    throw;
                }
            }
            catch (Exception ex)
            {
                return Json(new { message = "ERROR: " + ex.Message });
            }
        }

        public IActionResult ReportCategories()
        {
            return View("~/Views/MSaleAdmin/ReportCategories.cshtml");
        }


        [HttpGet]
        public IActionResult TareSetup()
        {
            return View("~/Views/MSaleAdmin/TareSetup.cshtml");
        }

        [HttpGet]
        public async Task<IActionResult> GetTares()
        {
            var list = new List<TareDto>();

            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            await cn.OpenAsync();

            var sql = "SELECT [Tare Name], [Tare Wt] FROM [Tare] ORDER BY [Tare Name]";

            using var cmd = new SqlCommand(sql, cn);
            using var rd = await cmd.ExecuteReaderAsync();

            while (await rd.ReadAsync())
            {
                list.Add(new TareDto
                {
                    TareName = rd["Tare Name"].ToString() ?? "",
                    TareWt = decimal.TryParse(rd["Tare Wt"].ToString(), out var wt) ? wt : 0
                });
            }

            return Json(list);
        }

        [HttpPost]
        public async Task<IActionResult> SaveTare([FromBody] TareDto item)
        {
            if (string.IsNullOrWhiteSpace(item.TareName))
                return Json(new { success = false, message = "Tare name is required." });

            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            await cn.OpenAsync();

            var sql = @"
IF EXISTS (SELECT 1 FROM [Tare] WHERE [Tare Name] = @TareName)
BEGIN
    UPDATE [Tare]
    SET [Tare Wt] = @TareWt
    WHERE [Tare Name] = @TareName
END
ELSE
BEGIN
    INSERT INTO [Tare] ([Tare Name], [Tare Wt])
    VALUES (@TareName, @TareWt)
END";

            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@TareName", item.TareName.Trim());
            cmd.Parameters.AddWithValue("@TareWt", item.TareWt);

            await cmd.ExecuteNonQueryAsync();

            return Json(new { success = true, message = "Saved." });
        }

        [HttpPost]
        public async Task<IActionResult> DeleteTare([FromBody] TareDto item)
        {
            if (string.IsNullOrWhiteSpace(item.TareName))
                return Json(new { success = false, message = "Nothing to delete." });

            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            await cn.OpenAsync();

            var sql = "DELETE FROM [Tare] WHERE [Tare Name] = @TareName";

            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@TareName", item.TareName.Trim());

            await cmd.ExecuteNonQueryAsync();

            return Json(new { success = true, message = "Deleted." });
        }

        [HttpGet]
        public IActionResult GetReportCategoryLists()
        {
            try
            {
                using var conn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                var departments = conn.Query<string>(@"
            SELECT [Report Name]
            FROM [dbo].[Report Categories]
            WHERE [Report Name] IS NOT NULL
            ORDER BY [Report Name]
        ").ToList();

                var masterDepartments = conn.Query<string>(@"
            SELECT [Master Name]
            FROM [dbo].[Master Departments]
            WHERE [Master Name] IS NOT NULL
            ORDER BY [Master Name]
        ").ToList();

                var timedEvents = conn.Query<string>(@"
            SELECT DISTINCT [Event Name]
            FROM [dbo].[Timed Pricing]
            WHERE [Event Name] IS NOT NULL
            ORDER BY [Event Name]
        ").ToList();

                // Leave empty until we know the real printer table
                var printers = conn.Query<string>(@"
    SELECT [Printer Name]
    FROM [dbo].[System Printers]
    WHERE [Printer Name] IS NOT NULL
    ORDER BY [Printer Name]
").ToList();




                return Json(new
                {
                    ok = true,
                    departments,
                    masterDepartments,
                    timedEvents,
                    printers
                });
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    ok = false,
                    error = ex.Message
                });
            }
        }

        [HttpGet]
        public IActionResult GetReportCategory(string reportName)
        {
            try
            {
                using var conn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                var row = conn.QueryFirstOrDefault(@"
            SELECT
                [Report Name] AS reportName,
                ISNULL([Flag], 0) AS flag,
                ISNULL([Record Number], 0) AS recordNumber,
                ISNULL([Remote 1], '') AS remote1,
                ISNULL([Remote 2], '') AS remote2,
                ISNULL([Remote 3], '') AS remote3,
                ISNULL([Remote 4], '') AS remote4,
                ISNULL([Remote 5], '') AS remote5,
                ISNULL([Remote 6], '') AS remote6,
                ISNULL([Master Name], 'None') AS masterName,
                ISNULL([Timed Event 1], 'None') AS timedEvent1,
                ISNULL([Timed Event 2], 'None') AS timedEvent2,
                ISNULL([Timed Event 3], 'None') AS timedEvent3,
                ISNULL([Timed Event 4], 'None') AS timedEvent4,
                ISNULL([Timed Event 5], 'None') AS timedEvent5,
                ISNULL([Department No], 0) AS departmentNo
            FROM [dbo].[Report Categories]
            WHERE [Report Name] = @reportName
        ", new { reportName });

                return Json(row ?? new { reportName });
            }
            catch (Exception ex)
            {
                return Json(new { error = ex.Message });
            }
        }
        [HttpGet]
        public IActionResult GetMealPeriods()
        {
            List<object> list = new List<object>();

            using (SqlConnection cn =
                new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"SELECT [Period Name],
                     [Period Time]
              FROM [Meal Period]
              ORDER BY [Period Name]", cn);

                SqlDataReader rd = cmd.ExecuteReader();

                while (rd.Read())
                {
                    list.Add(new
                    {
                        PeriodName = rd["Period Name"].ToString(),
                        PeriodTime = rd["Period Time"].ToString()
                    });
                }
            }

            return Json(list);
        }
        public IActionResult CategorySetup()
        {
            return View("~/Views/MSaleAdmin/CategorySetup.cshtml");
        }
        [HttpGet]
        public IActionResult GetCategory(string categoryName)
        {
            CategorySetupModel model = new CategorySetupModel();

            using (SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"SELECT [Category Name],
                     [Sort Level],
                     [Button Level],
                     [Mandatory],
                     [Disable Mods],
                     [Tip Exempt],
                     [Prompt],
                     [Meal Period],
                     [BitMap],
                     [Alias]
              FROM [Category Information]
              WHERE [Category Name] = @name", cn);

                cmd.Parameters.AddWithValue("@name", categoryName);

                SqlDataReader rd = cmd.ExecuteReader();

                if (rd.Read())
                {
                    model.CategoryName = rd["Category Name"].ToString();
                    model.SortLevel = rd["Sort Level"].ToString();
                    model.ButtonLevel = rd["Button Level"].ToString();
                    model.Mandatory = rd["Mandatory"].ToString();
                    model.DisableMods = rd["Disable Mods"].ToString();
                    model.TipExempt = rd["Tip Exempt"].ToString();
                    model.Prompt = rd["Prompt"].ToString();
                    model.MealPeriod = rd["Meal Period"].ToString();
                    model.BitMap = rd["BitMap"].ToString();
                    model.Alias = rd["Alias"].ToString();
                }
            }

            return Json(model);
        }
        [HttpPost]
        public IActionResult SaveCategory([FromBody] CategorySetupModel model)
        {
            using (SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand check = new SqlCommand(
                    @"SELECT COUNT(*)
              FROM [Category Information]
              WHERE [Category Name] = @name", cn);

                check.Parameters.AddWithValue("@name", model.CategoryName);

                int count = Convert.ToInt32(check.ExecuteScalar());

                if (count == 0)
                {
                    SqlCommand ins = new SqlCommand(
                        @"INSERT INTO [Category Information]
                  ([Category Name],
                   [Sort Level],
                   [Button Level],
                   [Mandatory],
                   [Disable Mods],
                   [Tip Exempt],
                   [Prompt],
                   [Meal Period],
                   [BitMap],
                   [Alias])
                  VALUES
                  (@CategoryName,
                   @SortLevel,
                   @ButtonLevel,
                   @Mandatory,
                   @DisableMods,
                   @TipExempt,
                   @Prompt,
                   @MealPeriod,
                   @BitMap,
                   @Alias)", cn);

                    AddCategoryParams(ins, model);
                    ins.ExecuteNonQuery();
                }
                else
                {
                    SqlCommand upd = new SqlCommand(
                        @"UPDATE [Category Information]
                  SET [Sort Level] = @SortLevel,
                      [Button Level] = @ButtonLevel,
                      [Mandatory] = @Mandatory,
                      [Disable Mods] = @DisableMods,
                      [Tip Exempt] = @TipExempt,
                      [Prompt] = @Prompt,
                      [Meal Period] = @MealPeriod,
                      [BitMap] = @BitMap,
                      [Alias] = @Alias
                  WHERE [Category Name] = @CategoryName", cn);

                    AddCategoryParams(upd, model);
                    upd.ExecuteNonQuery();
                }
            }

            return Json(new { success = true });
        }
        private void AddCategoryParams(SqlCommand cmd, CategorySetupModel model)
        {
            cmd.Parameters.AddWithValue("@CategoryName", model.CategoryName ?? "");
            cmd.Parameters.AddWithValue("@SortLevel", model.SortLevel ?? "");
            cmd.Parameters.AddWithValue("@ButtonLevel", model.ButtonLevel ?? "");
            cmd.Parameters.AddWithValue("@Mandatory", model.Mandatory ?? "No");
            cmd.Parameters.AddWithValue("@DisableMods", model.DisableMods ?? "No");
            cmd.Parameters.AddWithValue("@TipExempt", model.TipExempt ?? "No");
            cmd.Parameters.AddWithValue("@Prompt", model.Prompt ?? "");
            cmd.Parameters.AddWithValue("@MealPeriod", model.MealPeriod ?? "");
            cmd.Parameters.AddWithValue("@BitMap", model.BitMap ?? "");
            cmd.Parameters.AddWithValue("@Alias", model.Alias ?? "");
        }
        [HttpPost]
        public IActionResult DeleteCategory(string categoryName)
        {
            using (SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"DELETE FROM [Category Information]
              WHERE [Category Name] = @name", cn);

                cmd.Parameters.AddWithValue("@name", categoryName);
                cmd.ExecuteNonQuery();
            }

            return Json(new { success = true });
        }
        [HttpGet]
        public IActionResult GetMealPeriodNames()
        {
            List<string> list = new List<string>();

            using (SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"SELECT [Period Name]
              FROM [Meal Period]
              ORDER BY [Period Name]", cn);

                SqlDataReader rd = cmd.ExecuteReader();

                while (rd.Read())
                    list.Add(rd["Period Name"].ToString());
            }

            return Json(list);
        }
        [HttpGet]
        public IActionResult GetCategoryNames()
        {
            List<string> list = new List<string>();

            using (SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                SqlCommand cmd = new SqlCommand(
                    @"SELECT [Category Name]
              FROM [Category Information]
              ORDER BY [Category Name]", cn);

                SqlDataReader rd = cmd.ExecuteReader();

                while (rd.Read())
                    list.Add(rd["Category Name"].ToString());
            }

            return Json(list);
        }
        [HttpPost]
        public IActionResult SaveReportCategory([FromBody] ReportCategoryDto item)
        {
            if (item == null)
                return Json(new { message = "ERROR: No data received." });

            if (string.IsNullOrWhiteSpace(item.ReportName))
                return Json(new { message = "ERROR: Report name is required." });

            try
            {
                using var conn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                var exists = conn.ExecuteScalar<int>(@"
            SELECT COUNT(*)
            FROM [dbo].[Report Categories]
            WHERE [Report Name] = @ReportName
        ", item);

                if (exists > 0)
                {
                    conn.Execute(@"
                UPDATE [dbo].[Report Categories]
                SET
                    [Flag] = @Flag,
                    [Record Number] = @RecordNumber,
                    [Remote 1] = @Remote1,
                    [Remote 2] = @Remote2,
                    [Remote 3] = @Remote3,
                    [Remote 4] = @Remote4,
                    [Remote 5] = @Remote5,
                    [Remote 6] = @Remote6,
                    [Master Name] = @MasterName,
                    [Timed Event 1] = @TimedEvent1,
                    [Timed Event 2] = @TimedEvent2,
                    [Timed Event 3] = @TimedEvent3,
                    [Timed Event 4] = @TimedEvent4,
                    [Timed Event 5] = @TimedEvent5,
                    [Department No] = @DepartmentNo
                WHERE [Report Name] = @ReportName
            ", item);
                }
                else
                {
                    conn.Execute(@"
                INSERT INTO [dbo].[Report Categories]
                ([Report Name], [Flag], [Record Number],
                 [Remote 1], [Remote 2], [Remote 3], [Remote 4], [Remote 5], [Remote 6],
                 [Master Name],
                 [Timed Event 1], [Timed Event 2], [Timed Event 3], [Timed Event 4], [Timed Event 5],
                 [Department No])
                VALUES
                (@ReportName, @Flag, @RecordNumber,
                 @Remote1, @Remote2, @Remote3, @Remote4, @Remote5, @Remote6,
                 @MasterName,
                 @TimedEvent1, @TimedEvent2, @TimedEvent3, @TimedEvent4, @TimedEvent5,
                 @DepartmentNo)
            ", item);
                }

                return Json(new { message = "Saved" });
            }
            catch (Exception ex)
            {
                return Json(new { message = "ERROR: " + ex.Message });
            }
        }
        [HttpPost]
        public IActionResult DeleteReportCategory(string reportName)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            conn.Open();

            using var tran = conn.BeginTransaction();

            try
            {
                conn.Execute(@"
            DELETE FROM [Report Categories]
            WHERE [Report Name] = @reportName
        ", new { reportName }, tran);

                conn.Execute(@"
            DELETE FROM [NonSaleRevenue]
            WHERE [Name] = @reportName
        ", new { reportName }, tran);

                tran.Commit();

                return Json(new { message = "Deleted" });
            }
            catch (Exception ex)
            {
                tran.Rollback();
                return Json(new { message = "ERROR: " + ex.Message });
            }
        }

        public IActionResult TaxSetup()
        {
            return View();
        }

        [HttpPost]
        public IActionResult SaveTaxSetup([FromBody] TaxSetupSaveRequest data)
        {
            try
            {
                using (var conn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
                {
                    conn.Open();

                    conn.Execute("DELETE FROM [Sales Tax]");
                    conn.Execute("DELETE FROM [Special Tax Info]");

                    foreach (var t in data.SalesTaxes)
                    {
                        conn.Execute(@"
                    INSERT INTO [Sales Tax]
                    ([TaxTable], [Tax Name], [Tax Amount])
                    VALUES (@TaxTable, @TaxName, @TaxAmount)",
                            new
                            {
                                TaxTable = t.TaxTable,
                                TaxName = t.TaxName,
                                TaxAmount = t.TaxAmount
                            });
                    }

                    foreach (var t in data.SpecialTaxes)
                    {
                        conn.Execute(@"
                    INSERT INTO [Special Tax Info]
                    ([TaxTable], [Tax Name], [Special Tax], [Min Amount], [Max Amount], [Exemption])
                    VALUES (@TaxTable, @TaxName, @SpecialTax, @MinAmount, @MaxAmount, @Exemption)",
                            new
                            {
                                TaxTable = t.TaxTable,
                                TaxName = t.TaxName,
                                SpecialTax = t.SpecialTax,
                                MinAmount = t.MinTax,
                                MaxAmount = t.MaxTax,
                                Exemption = t.Exception
                            });
                    }
                }

                return Json(new { message = "Saved" });
            }
            catch (Exception ex)
            {
                return Json(new { message = "Error: " + ex.Message });
            }
        }
        [HttpPost]
        public IActionResult SaveTaxBreaks([FromBody] dynamic data)
        {
            int taxTable = (int)data.taxTable;

            using (var conn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                conn.Execute("DELETE FROM [Tax Breaks] WHERE TaxTable = @TaxTable",
                    new { TaxTable = taxTable });

                foreach (var r in data.rows)
                {
                    conn.Execute(@"
                INSERT INTO [Tax Breaks]
                (TaxTable, RowNumber, StartAmount, EndAmount, TaxAmount)
                VALUES (@TaxTable, @RowNumber, @StartAmount, @EndAmount, @TaxAmount)",
                        new
                        {
                            TaxTable = taxTable,
                            RowNumber = (int)r.rowNumber,
                            StartAmount = (decimal)r.startAmount,
                            EndAmount = (decimal)r.endAmount,
                            TaxAmount = (decimal)r.taxAmount
                        });
                }
            }

            return Json(new { message = "Saved" });
        }
        [HttpPost]
        public IActionResult SaveTaxBreaks([FromBody] TaxBreakSaveRequest data)
        {
            using var conn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection"));

            conn.Open();

            conn.Execute("DELETE FROM [Tax Breaks] WHERE [Tax Table] = @TaxTable",
                new { data.TaxTable });

            foreach (var r in data.Rows)
            {
                conn.Execute(@"
            INSERT INTO [Tax Breaks]
            ([Tax Table], [Row], [Start], [End], [Tax])
            VALUES (@TaxTable, @Row, @Start, @End, @Tax)",
                    new
                    {
                        TaxTable = data.TaxTable,
                        Row = r.RowNumber,
                        Start = r.StartAmount,
                        End = r.EndAmount,
                        Tax = r.TaxAmount
                    });
            }

            return Json(new { message = "Saved" });
        }

        [HttpGet]
        public IActionResult GetTaxSetup()
        {
            try
            {
                using var conn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                conn.Open();

                var salesTaxes = conn.Query<SalesTaxDto>(@"
            SELECT
                [TaxTable] AS TaxTable,
                [Tax Name] AS TaxName,
                [Tax Amount] AS TaxAmount
            FROM [Sales Tax]
            ORDER BY [TaxTable]
        ").ToList();

                var specialTaxes = conn.Query<SpecialTaxDto>(@"
            SELECT
                [TaxTable] AS TaxTable,
                [Tax Name] AS TaxName,
                [Special Tax] AS SpecialTax,
                [Min Amount] AS MinTax,
                [Max Amount] AS MaxTax,
                [Exemption] AS Exception
            FROM [Special Tax Info]
            ORDER BY [TaxTable]
        ").ToList();

                return Json(new
                {
                    salesTaxes,
                    specialTaxes
                });
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    error = ex.Message
                });
            }
        }

        [HttpGet]
        public IActionResult GetTaxBreaks(int taxTable)
        {
            try
            {
                using var conn = new SqlConnection(
                    _configuration.GetConnectionString("DefaultConnection"));

                conn.Open();

                var rows = conn.Query(@"
            SELECT
                [Tax Table] AS taxTable,
                [Row] AS rowNumber,
                [Start] AS startAmount,
                [End] AS endAmount,
                [Tax] AS taxAmount
            FROM [Tax Breaks]
            WHERE [Tax Table] = @taxTable
            ORDER BY [Row]
        ", new { taxTable }).ToList();

                return Json(rows);
            }
            catch (Exception ex)
            {
                return Json(new { error = ex.Message });
            }
        }
        [HttpPost]

        public IActionResult Save(DiscountEditModel model, int[] SelectedDays)
        {
            string cs = _configuration.GetConnectionString("DefaultConnection");

            using var cn = new SqlConnection(cs);
            cn.Open();

            string sql = @"
IF EXISTS(
    SELECT 1
    FROM[dbo].[Discounts]
    WHERE[Discount Name] = @name
)
BEGIN
    UPDATE[dbo].[Discounts]
            SET
        [Discount Type] = @type,
        [Discount Amount] = @amount,
        [Access Level] = @access,
        [Taxable] = @taxable,
        [Tax 1] = @tax1,
        [Tax 2] = @tax2,
        [Start Time] = @startTime,
        [End Time] = @endTime,
        [Expires] = @expires,
        [Allow Only One] = @allowOnlyOne,
        [After Print] = @afterPrint,
        [Auto Match] = @autoMatch,
        [Auto Deduct] = @autoDeduct,
        [Tracking] = @tracking,
        [Search Number] = @searchNumber
    WHERE[Discount Name] = @name
END
ELSE
BEGIN
    INSERT INTO[dbo].[Discounts]
    (
        [Discount Name],
        [Discount Type],
        [Discount Amount],
        [Access Level],
        [Taxable],
        [Tax 1],
        [Tax 2],
        [Start Time],
        [End Time],
        [Expires],
        [Allow Only One],
        [After Print],
        [Auto Match],
        [Auto Deduct],
        [Tracking],
        [Search Number]
    )
    VALUES
    (
        @name,
        @type,
        @amount,
        @access,
        @taxable,
        @tax1,
        @tax2,
        @startTime,
        @endTime,
        @expires,
        @allowOnlyOne,
        @afterPrint,
        @autoMatch,
        @autoDeduct,
        @tracking,
        @searchNumber
    )
END";
            using var cmd = new SqlCommand(sql, cn);

            cmd.Parameters.AddWithValue(
              "@name",
              model.DiscountName ?? "");

            cmd.Parameters.AddWithValue("@type", model.DiscountType ?? "");

            cmd.Parameters.AddWithValue(
              "@amount",
              model.DiscountAmount);

            cmd.Parameters.AddWithValue("@access", model.AccessLevel ?? "");
            cmd.Parameters.AddWithValue("@taxable", model.Taxable ? 1 : 0);

            cmd.Parameters.AddWithValue("@tax1", model.Tax1);
            cmd.Parameters.AddWithValue("@tax2", model.Tax2);
            cmd.Parameters.AddWithValue("@startTime", model.StartTime ?? "");
            cmd.Parameters.AddWithValue("@endTime", model.EndTime ?? "");
            cmd.Parameters.AddWithValue("@expires", model.Expires ?? "");
            cmd.Parameters.AddWithValue("@allowOnlyOne", model.AllowOnlyOne ? 1 : 0);
            cmd.Parameters.AddWithValue("@afterPrint", model.AfterPrint ? 1 : 0);
            cmd.Parameters.AddWithValue("@autoMatch", model.AutoMatch ? 1 : 0);
            cmd.Parameters.AddWithValue("@autoDeduct", model.AutoDeduct ? 1 : 0);
            cmd.Parameters.AddWithValue("@tracking", model.Tracking ? 1 : 0);
            cmd.Parameters.AddWithValue("@eventOnly", model.EventOnly ? 1 : 0);
            cmd.ExecuteNonQuery();

            string deleteSql =
"DELETE FROM [Active Days] WHERE [Discount Name] = @name";

            using (var del =
             new SqlCommand(deleteSql, cn))
            {
                del.Parameters.AddWithValue(
                  "@name",
                  model.DiscountName);

                del.ExecuteNonQuery();
            }

            if (SelectedDays != null)
            {
                foreach (var day in SelectedDays)
                {
                    string ins = @"
INSERT INTO [Active Days]
(
 [Day],
 [Pos],
 [Discount Name]
)
VALUES
(
 @day,
 @pos,
 @name
)";

                    using var cmdDay =
                       new SqlCommand(ins, cn);

                    string[] names =
                    {
    "Monday","Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday"
   };

                    cmdDay.Parameters.AddWithValue(
                      "@day",
                      names[day]);

                    cmdDay.Parameters.AddWithValue(
                      "@pos",
                       day);

                    cmdDay.Parameters.AddWithValue(
                      "@name",
                       model.DiscountName);

                    cmdDay.ExecuteNonQuery();
                }
            }

            return RedirectToAction("Index");
        }

        private List<DiscountTypeOption> GetDiscountTypes()
        {
            var list =
              new List<DiscountTypeOption>();

            string cs =
              _configuration.GetConnectionString(
               "DefaultConnection");

            using var cn =
              new SqlConnection(cs);

            cn.Open();

            string sql = @"
SELECT DiscountType,
       DisplayName
FROM DiscountTypes
WHERE isActive=1
ORDER BY GroupName, DisplayName";

            using var cmd =
              new SqlCommand(sql, cn);

            using var rdr =
              cmd.ExecuteReader();

            while (rdr.Read())
            {
                list.Add(
                 new DiscountTypeOption
                 {
                     DiscountType =
                      rdr["DiscountType"].ToString(),

                     DisplayName =
                      rdr["DisplayName"].ToString()
                 });
            }

            return list;
        }
        private List<DiscountCategoryOption>
GetDiscountCategories()
        {
            var list =
              new List<DiscountCategoryOption>();

            string cs =
              _configuration.GetConnectionString(
               "DefaultConnection");

            using var cn =
              new SqlConnection(cs);

            cn.Open();

            string sql = @"
SELECT CategoryName
FROM DiscountCategories
WHERE Active=1
ORDER BY SortOrder";

            using var cmd =
              new SqlCommand(sql, cn);

            using var rdr =
              cmd.ExecuteReader();

            while (rdr.Read())
            {
                list.Add(
                 new DiscountCategoryOption
                 {
                     CategoryName =
                    rdr["CategoryName"].ToString()
                 });
            }

            return list;
        }
        private List<SalesTaxOption> GetSalesTaxOptions()
        {
            var list = new List<SalesTaxOption>();
            string cs = _configuration.GetConnectionString("DefaultConnection");

            using var cn = new SqlConnection(cs);
            cn.Open();

            string sql = @"
SELECT [Tax Name], [Tax Amount], [TaxTable]
FROM [dbo].[Sales Tax]
ORDER BY [Tax Name]";

            using var cmd = new SqlCommand(sql, cn);
            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                list.Add(new SalesTaxOption
                {
                    TaxName = rdr["Tax Name"]?.ToString() ?? "",
                    TaxAmount = rdr["Tax Amount"] == DBNull.Value ? 0 : Convert.ToDecimal(rdr["Tax Amount"]),
                    TaxTable = rdr["TaxTable"] == DBNull.Value ? 0 : Convert.ToInt32(rdr["TaxTable"])
                });
            }

            return list;
        }
        private List<SpecialTaxOption> GetSpecialTaxOptions()
        {
            var list = new List<SpecialTaxOption>();
            string cs = _configuration.GetConnectionString("DefaultConnection");

            using var cn = new SqlConnection(cs);
            cn.Open();

            string sql = @"
SELECT [Special Tax], [Tax Name], [TaxTable]
FROM [dbo].[Special Tax Info]
ORDER BY [Tax Name], [Special Tax]";

            using var cmd = new SqlCommand(sql, cn);
            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                list.Add(new SpecialTaxOption
                {
                    SpecialTax = rdr["Special Tax"]?.ToString() ?? "",
                    TaxName = rdr["Tax Name"]?.ToString() ?? "",
                    TaxTable = rdr["TaxTable"]?.ToString() ?? ""
                });
            }

            return list;
        }
        private List<DiscountDayOption>
GetDiscountDays(
string discountName)
        {
            var allDays =
              new List<DiscountDayOption>
              {
        new(){Pos=0,Day="Monday"},
        new(){Pos=1,Day="Tuesday"},
        new(){Pos=2,Day="Wednesday"},
        new(){Pos=3,Day="Thursday"},
        new(){Pos=4,Day="Friday"},
        new(){Pos=5,Day="Saturday"},
        new(){Pos=6,Day="Sunday"}
              };

            if (string.IsNullOrWhiteSpace(
               discountName))
                return allDays;

            string cs =
              _configuration.GetConnectionString(
               "DefaultConnection");

            using var cn =
              new SqlConnection(cs);

            cn.Open();

            string sql = @"
SELECT Pos
FROM [Active Days]
WHERE [Discount Name]=@name";

            using var cmd =
              new SqlCommand(sql, cn);

            cmd.Parameters.AddWithValue(
              "@name",
              discountName);

            using var rdr =
              cmd.ExecuteReader();

            var selected =
              new HashSet<int>();

            while (rdr.Read())
            {
                selected.Add(
                  Convert.ToInt32(
                   rdr["Pos"]));
            }

            foreach (var d in allDays)
            {
                d.Selected =
                   selected.Contains(
                     d.Pos);
            }

            return allDays;
        }
        private static bool ToBool(object value)
        {
            if (value == null || value == DBNull.Value)
                return false;

            var s = value.ToString()?.Trim().ToLower() ?? "";

            return s == "1" || s == "true" || s == "yes" || s == "y";
        }


        //========================================
        // Load Menu Items and Modifiers
        //========================================
        [HttpGet]
        public JsonResult GetMenuAndModifiers()
        {
            var list = new List<object>();

            string cs =
                _configuration.GetConnectionString(
                    "DefaultConnection");

            using var cn =
                new SqlConnection(cs);

            cn.Open();

            string sql = @"
SELECT [Menu Item Name] AS Name
FROM [Menu Items]
WHERE ISNULL([Menu Item Name],'') <> ''

UNION

SELECT [Modifier Name] AS Name
FROM [Modifiers]
WHERE ISNULL([Modifier Name],'') <> ''

ORDER BY Name";

            using var cmd =
                new SqlCommand(sql, cn);

            using var rdr =
                cmd.ExecuteReader();

            while (rdr.Read())
            {
                list.Add(new
                {
                    name = rdr["Name"].ToString()
                });
            }

            return Json(list);
        }


        //========================================
        // Load Departments
        //========================================
        [HttpGet]
        public IActionResult GetDepartments()
        {
            var list = new List<string>();

            string cs = _configuration.GetConnectionString("DefaultConnection");

            using (SqlConnection cn = new SqlConnection(cs))
            {
                cn.Open();

                string sql = @"
            SELECT [Report Name]
            FROM [Report Categories]
            WHERE ISNULL([Report Name], '') <> ''
            ORDER BY [Report Name]";

                using (SqlCommand cmd = new SqlCommand(sql, cn))
                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        list.Add(dr["Report Name"].ToString().Trim());
                    }
                }
            }

            return Json(list);
        }
        //========================================
        // Load Discount Event/List Names
        //========================================        
        [HttpGet]
        public JsonResult GetDiscountEventNames()
        {
            var list = new List<string>();

            string cs =
                _configuration.GetConnectionString(
                    "DefaultConnection");

            using var cn =
                new SqlConnection(cs);

            cn.Open();

            string sql = @"
SELECT DISTINCT [Discount Name]
FROM [Discounts]
WHERE ISNULL([Discount Name],'') <> ''
ORDER BY [Discount Name]";

            using var cmd =
                new SqlCommand(sql, cn);

            using var rdr =
                cmd.ExecuteReader();

            while (rdr.Read())
            {
                list.Add(
                    rdr["Discount Name"].ToString()
                );
            }

            return Json(list);
        }


        [HttpGet]
        public IActionResult GetCouponAssignments(string couponName)
        {
            var list = new List<string>();

            if (string.IsNullOrWhiteSpace(couponName))
                return Json(list);

            string cs = _configuration.GetConnectionString("DefaultConnection");

            using (SqlConnection cn = new SqlConnection(cs))
            {
                cn.Open();

                string sql = @"
            SELECT [LinkToName]
            FROM [CouponAssignments]
            WHERE [CouponName] = @CouponName
            ORDER BY [LinkToName]";

                using (SqlCommand cmd = new SqlCommand(sql, cn))
                {
                    cmd.Parameters.AddWithValue("@CouponName", couponName);

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            list.Add(dr["LinkToName"].ToString().Trim());
                        }
                    }
                }
            }

            return Json(list);
        }

        [HttpPost]
        public JsonResult SaveCouponAssignments(
    [FromBody] CouponAssignmentSaveModel model)
        {
            try
            {
                string cs =
                    _configuration.GetConnectionString(
                        "DefaultConnection");

                using var cn =
                    new SqlConnection(cs);

                cn.Open();

                string deleteSql = @"
DELETE FROM [Coupon Assignments]
WHERE [Coupon Name] = @couponName";

                using (var del =
                    new SqlCommand(deleteSql, cn))
                {
                    del.Parameters.AddWithValue(
                        "@couponName",
                        model.CouponName ?? "");

                    del.ExecuteNonQuery();
                }

                if (model.LinkToNames != null)
                {
                    foreach (string item in model.LinkToNames)
                    {
                        string insertSql = @"
INSERT INTO [Coupon Assignments]
(
    [Coupon Name],
    [Link To Name]
)
VALUES
(
    @couponName,
    @linkToName
)";

                        using var cmd =
                            new SqlCommand(insertSql, cn);

                        cmd.Parameters.AddWithValue(
                            "@couponName",
                            model.CouponName ?? "");

                        cmd.Parameters.AddWithValue(
                            "@linkToName",
                            item ?? "");

                        cmd.ExecuteNonQuery();
                    }
                }

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    success = false,
                    error = ex.Message
                });
            }
        }
        [HttpPost]
        public JsonResult SaveDiscountEventAssignments(
    [FromBody] DiscountEventAssignmentSaveModel model)
        {
            try
            {
                string cs =
                    _configuration.GetConnectionString(
                        "DefaultConnection");

                using var cn =
                    new SqlConnection(cs);

                cn.Open();

                string deleteSql = @"
DELETE FROM [Discount Events]
WHERE [Event Name] = @eventName";

                using (var del =
                    new SqlCommand(deleteSql, cn))
                {
                    del.Parameters.AddWithValue(
                        "@eventName",
                        model.EventName ?? "");

                    del.ExecuteNonQuery();
                }

                if (model.DiscountNames != null)
                {
                    int pos = 0;

                    foreach (string discountName in model.DiscountNames)
                    {
                        string insertSql = @"
INSERT INTO [Discount Events]
(
    [Event Name],
    [POS],
    [Discount Name]
)
VALUES
(
    @eventName,
    @pos,
    @discountName
)";

                        using var cmd =
                            new SqlCommand(insertSql, cn);

                        cmd.Parameters.AddWithValue(
                            "@eventName",
                            model.EventName ?? "");

                        cmd.Parameters.AddWithValue(
                            "@pos",
                            pos);

                        cmd.Parameters.AddWithValue(
                            "@discountName",
                            discountName ?? "");

                        cmd.ExecuteNonQuery();

                        pos++;
                    }
                }

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    success = false,
                    error = ex.Message
                });
            }
        }
        //========================================
        // Load All Discounts
        //========================================
        [HttpGet]
        public JsonResult GetAllDiscountNames()
        {
            var list = new List<string>();

            string cs =
                _configuration.GetConnectionString("DefaultConnection");

            using var cn =
                new SqlConnection(cs);

            cn.Open();

            string sql = @"
SELECT [Discount Name]
FROM [Discounts]
WHERE ISNULL([Discount Name], '') <> ''
ORDER BY [Discount Name]";

            using var cmd =
                new SqlCommand(sql, cn);

            using var rdr =
                cmd.ExecuteReader();

            while (rdr.Read())
            {
                list.Add(
                    rdr["Discount Name"].ToString()
                );
            }

            return Json(list);
        }
        [HttpGet]
        public JsonResult GetDiscountEventAssignments(string eventName)
        {
            var list = new List<string>();

            string cs =
                _configuration.GetConnectionString("DefaultConnection");

            using var cn =
                new SqlConnection(cs);

            cn.Open();

            string sql = @"
SELECT [Discount Name]
FROM [Discount Events]
WHERE [Event Name] = @eventName
ORDER BY [POS]";

            using var cmd =
                new SqlCommand(sql, cn);

            cmd.Parameters.AddWithValue(
                "@eventName",
                eventName ?? "");

            using var rdr =
                cmd.ExecuteReader();

            while (rdr.Read())
            {
                list.Add(
                    rdr["Discount Name"].ToString()
                );
            }

            return Json(list);
        }
        [HttpGet]
        public IActionResult Nutritional()
        {
            return View("~/Views/MSaleAdmin/Nutritional.cshtml");
        }
        [HttpGet]
        public JsonResult GetNutritionalMenuNames()
        {
            var names = new List<string>();

            using (SqlConnection cn = new SqlConnection(
                _configuration.GetConnectionString("DefaultConnection")))
            {
                cn.Open();

                string sql = @"
            SELECT DISTINCT [Menu Item Name]
            FROM [Menu Items]
            WHERE ISNULL([Menu Item Name], '') <> ''
            ORDER BY [Menu Item Name]";

                using (SqlCommand cmd = new SqlCommand(sql, cn))
                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        names.Add(dr["Menu Item Name"].ToString());
                    }
                }
            }

            return Json(names);
        }


        [HttpPost]
        public IActionResult SaveTender([FromBody] TenderDto t)
        {
            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
    UPDATE [Credit Cards]
    SET
        [EMV] = @EMV,
        [Is Pin Pad] = @PinPad,
        [House Account] = @House,
        [Is Bowler] = @Bowler,
        [Is Debit] = @Debit,
        [Vendor Percentage] = @Vendor,
        [Is Cash] = @Cash,
        [Is House] = @Drawer,
        [Access Level] = @Access,
        [ConvenienceFeeAmt] = @FeeAmt,
        [Fee Type] = @FeeType
    WHERE [Card Name] = @Name";

            using SqlCommand cmd = new SqlCommand(sql, cn);

            cmd.Parameters.AddWithValue("@Name", t.TenderName);
            cmd.Parameters.AddWithValue("@EMV", t.UsesEMV);
            cmd.Parameters.AddWithValue("@PinPad", t.IsGiftEPay);
            cmd.Parameters.AddWithValue("@House", t.IsHouseAccount);
            cmd.Parameters.AddWithValue("@Bowler", t.IsKiosk);
            cmd.Parameters.AddWithValue("@Debit", t.IsInterCard);
            cmd.Parameters.AddWithValue("@Vendor", t.IsPropertyManagement);
            cmd.Parameters.AddWithValue("@Cash", t.IsCash);
            cmd.Parameters.AddWithValue("@Drawer", t.OpensDrawer);
            cmd.Parameters.AddWithValue("@Access", t.ManagerAccess);
            cmd.Parameters.AddWithValue("@FeeAmt", t.ConvenienceFeeAmt);
            cmd.Parameters.AddWithValue("@FeeType", t.FeeType ?? "");

            int rows = cmd.ExecuteNonQuery();

            if (rows == 0)
                return Json(new { success = false, message = "Tender was not found. Nothing saved." });

            return Json(new { success = true, message = "Saved" });
        }
        [HttpGet]
        public IActionResult GetTenders()
        {
            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        SELECT [Card Name]
        FROM [Credit Cards]
        ORDER BY [Card Name]";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            using SqlDataReader r = cmd.ExecuteReader();

            var list = new List<object>();

            while (r.Read())
            {
                list.Add(new
                {
                    tenderName = r["Card Name"] + ""
                });
            }

            return Json(list);
        }

        [HttpGet]
        public IActionResult GetTender(string name)
        {
            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        SELECT *
        FROM [Credit Cards]
        WHERE [Card Name] = @Name";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@Name", name);

            using SqlDataReader r = cmd.ExecuteReader();

            if (!r.Read())
                return NotFound();

            var tender = new
            {
                tenderName = r["Card Name"] + "",

                usesEMV = ToBool(r["EMV"]),
                isGiftEPay = ToBool(r["Is Pin Pad"]),
                isHouseAccount = ToBool(r["House Account"]),
                isKiosk = ToBool(r["Is Bowler"]),
                isInterCard = ToBool(r["Is Debit"]),
                isPropertyManagement = ToBool(r["Vendor Percentage"]),
                isCash = ToBool(r["Is Cash"]),
                opensDrawer = ToBool(r["Is House"]),
                managerAccess = ToBool(r["Access Level"]),

                convenienceFeeAmt = ToDecimal(r["ConvenienceFeeAmt"]),
                feeType = r["Fee Type"] + ""
            };

            return Json(tender);
        }


        private decimal ToDecimal(object value)
        {
            if (value == DBNull.Value || value == null)
                return 0;

            decimal.TryParse(value.ToString(), out decimal result);
            return result;
        }
        public IActionResult StorePaidOuts()
        {
            return View("~/Views/MSaleAdmin/StorePaidOuts.cshtml");
        }

        public class StorePaidOutDto
        {
            public string PaidOutName { get; set; } = "";
        }

        [HttpGet]
        public IActionResult GetStorePaidOuts()
        {
            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        SELECT [Paid Out Name]
        FROM [StorePaidOuts]
        ORDER BY [Paid Out Name]";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            using SqlDataReader r = cmd.ExecuteReader();

            var list = new List<StorePaidOutDto>();

            while (r.Read())
            {
                list.Add(new StorePaidOutDto
                {
                    PaidOutName = r["Paid Out Name"] + ""
                });
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveStorePaidOut([FromBody] StorePaidOutDto item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.PaidOutName))
                return Json(new { success = false, message = "Enter a paid out name." });

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        IF EXISTS (
            SELECT 1 FROM [StorePaidOuts]
            WHERE [Paid Out Name] = @PaidOutName
        )
        BEGIN
            UPDATE [StorePaidOuts]
            SET [Paid Out Name] = @PaidOutName
            WHERE [Paid Out Name] = @PaidOutName
        END
        ELSE
        BEGIN
            INSERT INTO [StorePaidOuts] ([Paid Out Name])
            VALUES (@PaidOutName)
        END";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@PaidOutName", item.PaidOutName.Trim());
            cmd.ExecuteNonQuery();

            return Json(new { success = true, message = "Saved" });
        }
        public IActionResult CouponReasons()
        {
            return View("~/Views/MSaleAdmin/CouponReasons.cshtml");
        }

        public class CouponReasonDto
        {
            public string Name { get; set; } = "";
            public string OriginalName { get; set; } = "";
        }

        [HttpGet]
        public IActionResult GetCouponReasons()
        {
            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        SELECT [Name]
        FROM [Coupon Reason]
        ORDER BY [Name]";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            using SqlDataReader r = cmd.ExecuteReader();

            var list = new List<CouponReasonDto>();

            while (r.Read())
            {
                list.Add(new CouponReasonDto
                {
                    Name = r["Name"] + ""
                });
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveCouponReason([FromBody] CouponReasonDto item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.Name))
                return Json(new { success = false, message = "Enter a reason name." });

            string newName = item.Name.Trim();
            string oldName = string.IsNullOrWhiteSpace(item.OriginalName)
                ? newName
                : item.OriginalName.Trim();

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        IF EXISTS (
            SELECT 1 FROM [Coupon Reason]
            WHERE [Name] = @OldName
        )
        BEGIN
            UPDATE [Coupon Reason]
            SET [Name] = @NewName
            WHERE [Name] = @OldName
        END
        ELSE
        BEGIN
            INSERT INTO [Coupon Reason] ([Name])
            VALUES (@NewName)
        END";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@OldName", oldName);
            cmd.Parameters.AddWithValue("@NewName", newName);
            cmd.ExecuteNonQuery();

            return Json(new { success = true, message = "Saved" });
        }

        [HttpPost]
        public IActionResult DeleteCouponReason([FromBody] CouponReasonDto item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.Name))
                return Json(new { success = false, message = "Select a reason to delete." });

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        DELETE FROM [Coupon Reason]
        WHERE [Name] = @Name";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@Name", item.Name.Trim());

            int rows = cmd.ExecuteNonQuery();

            return Json(new
            {
                success = rows > 0,
                message = rows > 0 ? "Deleted" : "Reason not found."
            });
        }
        public IActionResult ModifierPrefixes()
        {
            return View("~/Views/MSaleAdmin/ModifierPrefixes.cshtml");
        }

        public class ModifierPrefixDto
        {
            public string PrefixName { get; set; } = "";
            public string OriginalPrefixName { get; set; } = "";
            public decimal Upcharge { get; set; }
        }

        [HttpGet]
        public IActionResult GetModifierPrefixes()
        {
            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        SELECT [Prefix Name], [Upcharge]
        FROM [Modifier Prefix]
        ORDER BY [Prefix Name]";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            using SqlDataReader r = cmd.ExecuteReader();

            var list = new List<ModifierPrefixDto>();

            while (r.Read())
            {
                list.Add(new ModifierPrefixDto
                {
                    PrefixName = r["Prefix Name"] + "",
                    Upcharge = ToDecimal(r["Upcharge"])
                });
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveModifierPrefix([FromBody] ModifierPrefixDto item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.PrefixName))
                return Json(new { success = false, message = "Enter a prefix name." });

            string newName = item.PrefixName.Trim();
            string oldName = string.IsNullOrWhiteSpace(item.OriginalPrefixName)
                ? newName
                : item.OriginalPrefixName.Trim();

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        IF EXISTS (
            SELECT 1 FROM [Modifier Prefix]
            WHERE [Prefix Name] = @OldName
        )
        BEGIN
            UPDATE [Modifier Prefix]
            SET 
                [Prefix Name] = @NewName,
                [Upcharge] = @Upcharge
            WHERE [Prefix Name] = @OldName
        END
        ELSE
        BEGIN
            INSERT INTO [Modifier Prefix] ([Prefix Name], [Upcharge])
            VALUES (@NewName, @Upcharge)
        END";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@OldName", oldName);
            cmd.Parameters.AddWithValue("@NewName", newName);
            cmd.Parameters.AddWithValue("@Upcharge", item.Upcharge);
            cmd.ExecuteNonQuery();

            return Json(new { success = true, message = "Saved" });
        }

        [HttpPost]
        public IActionResult DeleteModifierPrefix([FromBody] ModifierPrefixDto item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.PrefixName))
                return Json(new { success = false, message = "Select a prefix to delete." });

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        DELETE FROM [Modifier Prefix]
        WHERE [Prefix Name] = @PrefixName";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@PrefixName", item.PrefixName.Trim());

            int rows = cmd.ExecuteNonQuery();

            return Json(new
            {
                success = rows > 0,
                message = rows > 0 ? "Deleted" : "Prefix not found."
            });
        }
        public IActionResult SpecialRequests()
        {
            return View("~/Views/MSaleAdmin/SpecialRequests.cshtml");
        }

        public class SpecialRequestDto
        {
            public string Name { get; set; } = "";
            public string OriginalName { get; set; } = "";
            public decimal Amount { get; set; }
        }

        [HttpGet]
        public IActionResult GetSpecialRequests()
        {
            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        SELECT [Request Name], [Amount]
        FROM [Special Requests]
        ORDER BY [Request Name]";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            using SqlDataReader r = cmd.ExecuteReader();

            var list = new List<SpecialRequestDto>();

            while (r.Read())
            {
                list.Add(new SpecialRequestDto
                {
                    Name = r["Request Name"] + "",
                    Amount = ToDecimal(r["Amount"])
                });
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveSpecialRequest([FromBody] SpecialRequestDto item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.Name))
                return Json(new { success = false, message = "Enter a request name." });

            string newName = item.Name.Trim();
            string oldName = string.IsNullOrWhiteSpace(item.OriginalName)
                ? newName
                : item.OriginalName.Trim();

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        IF EXISTS (
            SELECT 1 FROM [Special Requests]
            WHERE [Request Name] = @OldName
        )
        BEGIN
            UPDATE [Special Requests]
            SET
                [Request Name] = @NewName,
                [Amount] = @Amount
            WHERE [Request Name] = @OldName
        END
        ELSE
        BEGIN
            INSERT INTO [Special Requests]
            (
                [Request Name],
                [Amount]
            )
            VALUES
            (
                @NewName,
                @Amount
            )
        END";

            using SqlCommand cmd = new SqlCommand(sql, cn);

            cmd.Parameters.AddWithValue("@OldName", oldName);
            cmd.Parameters.AddWithValue("@NewName", newName);
            cmd.Parameters.AddWithValue("@Amount", item.Amount);

            cmd.ExecuteNonQuery();

            return Json(new { success = true, message = "Saved" });
        }

        [HttpPost]
        public IActionResult DeleteSpecialRequest([FromBody] SpecialRequestDto item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.Name))
                return Json(new { success = false, message = "Select a request to delete." });

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        DELETE FROM [Special Requests]
        WHERE [Request Name] = @Name";

            using SqlCommand cmd = new SqlCommand(sql, cn);

            cmd.Parameters.AddWithValue("@Name", item.Name.Trim());

            int rows = cmd.ExecuteNonQuery();

            return Json(new
            {
                success = rows > 0,
                message = rows > 0 ? "Deleted" : "Request not found."
            });
        }
        public IActionResult VoidReasons()
        {
            return View("~/Views/MSaleAdmin/VoidReasons.cshtml");
        }

        public class VoidReasonDto
        {
            public string VoidName { get; set; } = "";
            public string OriginalVoidName { get; set; } = "";
            public bool Inventory { get; set; }
        }

        [HttpGet]
        public IActionResult GetVoidReasons()
        {
            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        SELECT [Void Name], [Inventory]
        FROM [Void Reasons]
        ORDER BY [Void Name]";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            using SqlDataReader r = cmd.ExecuteReader();

            var list = new List<VoidReasonDto>();

            while (r.Read())
            {
                list.Add(new VoidReasonDto
                {
                    VoidName = r["Void Name"] + "",
                    Inventory = ToBool(r["Inventory"])
                });
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveVoidReason([FromBody] VoidReasonDto item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.VoidName))
                return Json(new { success = false, message = "Enter a void reason." });

            string newName = item.VoidName.Trim();

            string oldName = string.IsNullOrWhiteSpace(item.OriginalVoidName)
                ? newName
                : item.OriginalVoidName.Trim();

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        IF EXISTS (
            SELECT 1 FROM [Void Reasons]
            WHERE [Void Name] = @OldName
        )
        BEGIN
            UPDATE [Void Reasons]
            SET
                [Void Name] = @NewName,
                [Inventory] = @Inventory
            WHERE [Void Name] = @OldName
        END
        ELSE
        BEGIN
            INSERT INTO [Void Reasons]
            (
                [Void Name],
                [Inventory]
            )
            VALUES
            (
                @NewName,
                @Inventory
            )
        END";

            using SqlCommand cmd = new SqlCommand(sql, cn);

            cmd.Parameters.AddWithValue("@OldName", oldName);
            cmd.Parameters.AddWithValue("@NewName", newName);
            cmd.Parameters.AddWithValue("@Inventory", item.Inventory ? 1 : 0);

            cmd.ExecuteNonQuery();

            return Json(new { success = true, message = "Saved" });
        }

        [HttpPost]
        public IActionResult DeleteVoidReason([FromBody] VoidReasonDto item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.VoidName))
                return Json(new { success = false, message = "Select a void reason to delete." });

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        DELETE FROM [Void Reasons]
        WHERE [Void Name] = @VoidName";

            using SqlCommand cmd = new SqlCommand(sql, cn);

            cmd.Parameters.AddWithValue("@VoidName", item.VoidName.Trim());

            int rows = cmd.ExecuteNonQuery();

            return Json(new
            {
                success = rows > 0,
                message = rows > 0 ? "Deleted" : "Void reason not found."
            });
        }
        public IActionResult JobTitles()
        {
            return View("~/Views/MSaleAdmin/JobTitles.cshtml");
        }

        public class JobTitleDto
        {
            public string JobTitle { get; set; } = "";
            public string OriginalJobTitle { get; set; } = "";
            public bool TippedPosition { get; set; }
            public decimal TipPoolPercent { get; set; }
            public string JobCodeNumber { get; set; } = "";
            public bool Drawer { get; set; }
        }

        [HttpGet]
        public IActionResult GetJobTitles()
        {
            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        SELECT [Job Title], [Tipped Position], [Tip Pool Percent],
               [Job Code Number], [Drawer]
        FROM [Job Titles]
        ORDER BY [Job Title]";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            using SqlDataReader r = cmd.ExecuteReader();

            var list = new List<JobTitleDto>();

            while (r.Read())
            {
                list.Add(new JobTitleDto
                {
                    JobTitle = r["Job Title"] + "",
                    TippedPosition = ToBool(r["Tipped Position"]),
                    TipPoolPercent = ToDecimal(r["Tip Pool Percent"]),
                    JobCodeNumber = r["Job Code Number"] + "",
                    Drawer = ToBool(r["Drawer"])
                });
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveJobTitle([FromBody] JobTitleDto item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.JobTitle))
                return Json(new { success = false, message = "Enter a job title." });

            string newName = item.JobTitle.Trim();
            string oldName = string.IsNullOrWhiteSpace(item.OriginalJobTitle)
                ? newName
                : item.OriginalJobTitle.Trim();

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        IF EXISTS (
            SELECT 1 FROM [Job Titles]
            WHERE [Job Title] = @OldName
        )
        BEGIN
            UPDATE [Job Titles]
            SET
                [Job Title] = @NewName,
                [Tipped Position] = @TippedPosition,
                [Tip Pool Percent] = @TipPoolPercent,
                [Job Code Number] = @JobCodeNumber,
                [Drawer] = @Drawer
            WHERE [Job Title] = @OldName
        END
        ELSE
        BEGIN
            INSERT INTO [Job Titles]
            (
                [Job Title],
                [Tipped Position],
                [Tip Pool Percent],
                [Job Code Number],
                [Drawer]
            )
            VALUES
            (
                @NewName,
                @TippedPosition,
                @TipPoolPercent,
                @JobCodeNumber,
                @Drawer
            )
        END";

            using SqlCommand cmd = new SqlCommand(sql, cn);

            cmd.Parameters.AddWithValue("@OldName", oldName);
            cmd.Parameters.AddWithValue("@NewName", newName);
            cmd.Parameters.AddWithValue("@TippedPosition", item.TippedPosition ? 1 : 0);
            cmd.Parameters.AddWithValue("@TipPoolPercent", item.TipPoolPercent);
            cmd.Parameters.AddWithValue("@JobCodeNumber", item.JobCodeNumber ?? "");
            cmd.Parameters.AddWithValue("@Drawer", item.Drawer ? 1 : 0);

            cmd.ExecuteNonQuery();

            return Json(new { success = true, message = "Saved" });
        }

        [HttpPost]
        public IActionResult DeleteJobTitle([FromBody] JobTitleDto item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.JobTitle))
                return Json(new { success = false, message = "Select a job title to delete." });

            using SqlConnection cn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Open();

            string sql = @"
        DELETE FROM [Job Titles]
        WHERE [Job Title] = @JobTitle";

            using SqlCommand cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@JobTitle", item.JobTitle.Trim());

            int rows = cmd.ExecuteNonQuery();

            return Json(new
            {
                success = rows > 0,
                message = rows > 0 ? "Deleted" : "Job title not found."
            });
        }
        public class RefundReasonModel
        {
            public string RefundName { get; set; }
        }

        public class SaveRefundReasonModel
        {
            public string OldName { get; set; }
            public string RefundName { get; set; }
        }

        [HttpGet]
        public IActionResult RefundReasons()
        {
            return View("~/Views/MSaleAdmin/RefundReasons.cshtml");
        }

        [HttpGet]
        public JsonResult GetRefundReasons()
        {
            List<object> list = new List<object>();

            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql =
                    @"SELECT [Refund Name]
              FROM [Refund Reason]
              ORDER BY [Refund Name]";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        list.Add(new
                        {
                            refundName = rdr["Refund Name"].ToString()
                        });
                    }
                }
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveRefundReason(
            [FromBody] SaveRefundReasonModel model)
        {
            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                if (string.IsNullOrEmpty(model.OldName))
                {
                    string insertSql =
                        @"INSERT INTO [Refund Reason]
                  ([Refund Name])
                  VALUES
                  (@RefundName)";

                    using (SqlCommand cmd =
                        new SqlCommand(insertSql, conn))
                    {
                        cmd.Parameters.AddWithValue("@RefundName",
                            model.RefundName);

                        cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    string updateSql =
                        @"UPDATE [Refund Reason]
                  SET [Refund Name] = @RefundName
                  WHERE [Refund Name] = @OldName";

                    using (SqlCommand cmd =
                        new SqlCommand(updateSql, conn))
                    {
                        cmd.Parameters.AddWithValue("@RefundName",
                            model.RefundName);

                        cmd.Parameters.AddWithValue("@OldName",
                            model.OldName);

                        cmd.ExecuteNonQuery();
                    }
                }
            }

            return Content("Saved Successfully");
        }

        [HttpPost]
        public IActionResult DeleteRefundReason(
            [FromBody] RefundReasonModel model)
        {
            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql =
                    @"DELETE FROM [Refund Reason]
              WHERE [Refund Name] = @RefundName";

                using (SqlCommand cmd =
                    new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@RefundName",
                        model.RefundName);

                    cmd.ExecuteNonQuery();
                }
            }

            return Content("Deleted Successfully");
        }

        public class SavePreferenceModel
        {
            public string OldName { get; set; }
            public string Name { get; set; }
        }


        [HttpGet]
        public IActionResult Preferences()
        {
            return View("~/Views/MSaleAdmin/Preferences.cshtml");
        }

        [HttpGet]
        public JsonResult GetPreferences()
        {
            List<object> list = new List<object>();

            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql =
                    @"SELECT [Name]
              FROM [Preferences]
              ORDER BY [Name]";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        list.Add(new
                        {
                            name = rdr["Name"].ToString()
                        });
                    }
                }
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SavePreference(
            [FromBody] SavePreferenceModel model)
        {
            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                if (string.IsNullOrEmpty(model.OldName))
                {
                    string sql =
                        @"INSERT INTO [Preferences]
                  ([Name])
                  VALUES
                  (@Name)";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@Name", model.Name);
                        cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    string sql =
                        @"UPDATE [Preferences]
                  SET [Name] = @Name
                  WHERE [Name] = @OldName";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@Name", model.Name);
                        cmd.Parameters.AddWithValue("@OldName", model.OldName);
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            return Content("Saved Successfully");
        }
        public class StorePaidInModel
        {
            public int Id { get; set; }
            public string PaidInName { get; set; }
        }

        [HttpGet]
        public IActionResult StorePaidIn()
        {
            return View("~/Views/MSaleAdmin/StorePaidIn.cshtml");
        }

        [HttpGet]
        public JsonResult GetStorePaidIns()
        {
            var list = new List<StorePaidInModel>();

            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql =
                    @"SELECT [Id], [PaidInName]
              FROM [StorePaidIn]
              ORDER BY [PaidInName]";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        list.Add(new StorePaidInModel
                        {
                            Id = Convert.ToInt32(rdr["Id"]),
                            PaidInName = rdr["PaidInName"].ToString()
                        });
                    }
                }
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveStorePaidIn([FromBody] StorePaidInModel model)
        {
            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                if (model.Id == 0)
                {
                    string sql =
                        @"INSERT INTO [StorePaidIn]
                  ([PaidInName])
                  VALUES
                  (@PaidInName)";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@PaidInName", model.PaidInName);
                        cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    string sql =
                        @"UPDATE [StorePaidIn]
                  SET [PaidInName] = @PaidInName
                  WHERE [Id] = @Id";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@PaidInName", model.PaidInName);
                        cmd.Parameters.AddWithValue("@Id", model.Id);
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            return Content("Saved Successfully");
        }

        [HttpPost]
        public IActionResult DeleteStorePaidIn([FromBody] StorePaidInModel model)
        {
            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql =
                    @"DELETE FROM [StorePaidIn]
              WHERE [Id] = @Id";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@Id", model.Id);
                    cmd.ExecuteNonQuery();
                }
            }

            return Content("Deleted Successfully");
        }
        [HttpGet]
        public IActionResult Dashboard()
        {
            return View("~/Views/MSaleAdmin/Dashboard.cshtml");
        }
        [HttpGet]
        public IActionResult Employees()
        {
            return View("~/Views/MSaleAdmin/Employees.cshtml");
        }

        [HttpGet]
        public JsonResult GetEmployees(int terminated = 0)
        {
            List<object> list = new List<object>();

            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql =
                    @"SELECT
                [First Name],
                [Middle Name],
                [Last Name],
                [Address],
                [City],
                [State],
                [Zip],
                [Access Level],
                [Access Number],
                [Home Phone],
                [Phone # 2],
                [Language],
                [Employee Id],
                [Company Id],
                [Hire Date],
                [DOB],
                [Nick Name],
                [Terminated]
              FROM [General Data]
              WHERE ISNULL([Terminated], 0) = @Terminated
              ORDER BY [Last Name], [First Name]";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@Terminated", terminated);

                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            list.Add(new
                            {
                                firstName = rdr["First Name"].ToString(),
                                middleName = rdr["Middle Name"].ToString(),
                                lastName = rdr["Last Name"].ToString(),
                                address = rdr["Address"].ToString(),
                                city = rdr["City"].ToString(),
                                state = rdr["State"].ToString(),
                                zip = rdr["Zip"].ToString(),
                                accessLevel = rdr["Access Level"].ToString(),
                                accessNumber = rdr["Access Number"].ToString(),
                                homePhone = rdr["Home Phone"].ToString(),
                                phone2 = rdr["Phone # 2"].ToString(),
                                language = rdr["Language"].ToString(),
                                employeeId = rdr["Employee Id"].ToString(),
                                companyId = rdr["Company Id"].ToString(),
                                hireDate = FormatDateForInput(rdr["Hire Date"]),
                                dob = FormatDateForInput(rdr["DOB"]),
                                nickName = rdr["Nick Name"].ToString(),
                                terminated = Convert.ToInt32(rdr["Terminated"] == DBNull.Value ? 0 : rdr["Terminated"]) != 0
                            });
                        }
                    }
                }
            }

            return Json(list);
        }
        [HttpPost]
        public IActionResult SaveEmployee([FromBody] EmployeeAdminModel model)
        {
            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                bool isNew =
                    string.IsNullOrWhiteSpace(model.OriginalFirstName) ||
                    string.IsNullOrWhiteSpace(model.OriginalLastName);

                if (isNew)
                {
                    string insertSql =
                        @"INSERT INTO [General Data]
                  (
                    [First Name],
                    [Middle Name],
                    [Last Name],
                    [Address],
                    [City],
                    [State],
                    [Zip],
                    [Access Level],
                    [Access Number],
                    [Home Phone],
                    [Phone # 2],
                    [Language],
                    [Employee Id],
                    [Company Id],
                    [Hire Date],
                    [DOB],
                    [Nick Name],
                    [Terminated],
                    [Pager #],
                    [Sex],
                    [Marital Status],
                    [Ethnic],
                    [Federal Exemption],
                    [State Exemption],
                    [SSN],
                    [FP]
                  )
                  VALUES
                  (
                    @FirstName,
                    @MiddleName,
                    @LastName,
                    @Address,
                    @City,
                    @State,
                    @Zip,
                    @AccessLevel,
                    @AccessNumber,
                    @HomePhone,
                    @Phone2,
                    @Language,
                    @EmployeeId,
                    @CompanyId,
                    @HireDate,
                    @Dob,
                    @NickName,
                    @Terminated,
                    '',
                    '',
                    '',
                    '',
                    0,
                    0,
                    '',
                    ''
                  )";

                    using (SqlCommand cmd = new SqlCommand(insertSql, conn))
                    {
                        AddEmployeeParams(cmd, model);
                        cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    string updateSql =
                        @"UPDATE [General Data]
                  SET
                    [First Name] = @FirstName,
                    [Middle Name] = @MiddleName,
                    [Last Name] = @LastName,
                    [Address] = @Address,
                    [City] = @City,
                    [State] = @State,
                    [Zip] = @Zip,
                    [Access Level] = @AccessLevel,
                    [Access Number] = @AccessNumber,
                    [Home Phone] = @HomePhone,
                    [Phone # 2] = @Phone2,
                    [Language] = @Language,
                    [Employee Id] = @EmployeeId,
                    [Company Id] = @CompanyId,
                    [Hire Date] = @HireDate,
                    [DOB] = @Dob,
                    [Nick Name] = @NickName,
                    [Terminated] = @Terminated
                  WHERE [First Name] = @OriginalFirstName
                    AND [Last Name] = @OriginalLastName";

                    using (SqlCommand cmd = new SqlCommand(updateSql, conn))
                    {
                        AddEmployeeParams(cmd, model);
                        cmd.Parameters.AddWithValue("@OriginalFirstName", model.OriginalFirstName ?? "");
                        cmd.Parameters.AddWithValue("@OriginalLastName", model.OriginalLastName ?? "");
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            return Content("Employee Saved Successfully");
        }
        private void AddEmployeeParams(SqlCommand cmd, EmployeeAdminModel model)
        {
            cmd.Parameters.AddWithValue("@FirstName", model.FirstName ?? "");
            cmd.Parameters.AddWithValue("@MiddleName", model.MiddleName ?? "");
            cmd.Parameters.AddWithValue("@LastName", model.LastName ?? "");
            cmd.Parameters.AddWithValue("@Address", model.Address ?? "");
            cmd.Parameters.AddWithValue("@City", model.City ?? "");
            cmd.Parameters.AddWithValue("@State", model.State ?? "");
            cmd.Parameters.AddWithValue("@Zip", model.Zip ?? "");
            cmd.Parameters.AddWithValue("@AccessLevel", model.AccessLevel ?? "");
            cmd.Parameters.AddWithValue("@AccessNumber", model.AccessNumber ?? "");
            cmd.Parameters.AddWithValue("@HomePhone", model.HomePhone ?? "");
            cmd.Parameters.AddWithValue("@Phone2", model.Phone2 ?? "");
            cmd.Parameters.AddWithValue("@Language", model.Language ?? "");
            cmd.Parameters.AddWithValue("@EmployeeId", model.EmployeeId ?? "");
            cmd.Parameters.AddWithValue("@CompanyId", model.CompanyId ?? "");
            cmd.Parameters.AddWithValue("@HireDate", model.HireDate ?? "");
            cmd.Parameters.AddWithValue("@Dob", model.Dob ?? "");
            cmd.Parameters.AddWithValue("@NickName", model.NickName ?? "");
            cmd.Parameters.AddWithValue("@Terminated", model.Terminated ? 1 : 0);
        }
        [HttpPost]
        public IActionResult DeleteEmployee([FromBody] EmployeeAdminModel model)
        {
            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql =
                    @"DELETE FROM [General Data]
              WHERE [First Name] = @FirstName
                AND [Last Name] = @LastName";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@FirstName", model.OriginalFirstName ?? "");
                    cmd.Parameters.AddWithValue("@LastName", model.OriginalLastName ?? "");
                    cmd.ExecuteNonQuery();
                }
            }

            return Content("Employee Removed Successfully");
        }
        private string FormatDateForInput(object value)
        {
            if (value == null || value == DBNull.Value)
                return "";

            if (DateTime.TryParse(value.ToString(), out DateTime d))
                return d.ToString("yyyy-MM-dd");

            return "";
        }
        public class SuggestedTipModel
        {
            public string Name { get; set; }
            public decimal Amount { get; set; }
        }

        [HttpGet]
        public IActionResult SuggestedTips()
        {
            return View("~/Views/MSaleAdmin/SuggestedTips.cshtml");
        }

        [HttpGet]
        public JsonResult GetSuggestedTips()
        {
            List<SuggestedTipModel> list = new List<SuggestedTipModel>();

            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql =
                    @"SELECT [Name], [Amount]
              FROM [SuggestedTip]
              ORDER BY [Amount]";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        list.Add(new SuggestedTipModel
                        {
                            Name = rdr["Name"].ToString(),
                            Amount = Convert.ToDecimal(rdr["Amount"])
                        });
                    }
                }
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveSuggestedTips([FromBody] List<SuggestedTipModel> tips)
        {
            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                using (SqlTransaction tran = conn.BeginTransaction())
                {
                    try
                    {
                        string deleteSql = @"DELETE FROM [SuggestedTip]";

                        using (SqlCommand cmd = new SqlCommand(deleteSql, conn, tran))
                        {
                            cmd.ExecuteNonQuery();
                        }

                        string insertSql =
                            @"INSERT INTO [SuggestedTip]
                      ([Name], [Amount])
                      VALUES
                      (@Name, @Amount)";

                        foreach (var tip in tips)
                        {
                            using (SqlCommand cmd = new SqlCommand(insertSql, conn, tran))
                            {
                                cmd.Parameters.AddWithValue("@Name", tip.Name ?? "");
                                cmd.Parameters.AddWithValue("@Amount", tip.Amount);
                                cmd.ExecuteNonQuery();
                            }
                        }

                        tran.Commit();
                    }
                    catch
                    {
                        tran.Rollback();
                        throw;
                    }
                }
            }

            return Json(new { message = "Suggested Tips Saved" });
        }

        [HttpPost]
        public IActionResult DeleteSuggestedTips()
        {
            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql = @"DELETE FROM [SuggestedTip]";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }

            return Json(new { message = "Suggested Tips Deleted" });
        }
        public class DiscountAssignmentSaveRequest
        {
            public string EventName { get; set; }
            public List<DiscountAssignmentItem> Discounts { get; set; }
        }

        public class DiscountAssignmentItem
        {
            public int Pos { get; set; }
            public string DiscountName { get; set; }
        }

        [HttpGet]
        public IActionResult DiscountEventAssignments()
        {
            return View("~/Views/MSaleAdmin/DiscountEventAssignments.cshtml");
        }

        [HttpGet]
        public JsonResult GetDiscountAssignmentDiscounts()
        {
            List<string> list = new List<string>();

            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql =
                    @"SELECT [Discount Name]
              FROM [Discounts]
              ORDER BY [Discount Name]";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        list.Add(rdr["Discount Name"].ToString());
                    }
                }
            }

            return Json(list);
        }

        [HttpGet]
        public JsonResult GetDiscountAssignmentEvents()
        {
            List<string> list = new List<string>();

            string[] fixedEvents =
            {
        "Delivery",
        "Employee Discount",
        "Automatic Gratuity",
        "Service Charge",
        "NetEpayLoyalty",
        "Processing Fee",
        "Automatic Service Charge",
        "Cash Discount"
    };

            foreach (string e in fixedEvents)
                AddUnique(list, e);

            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql =
                    @"SELECT DISTINCT [Event Name]
              FROM [Discount Events]
              WHERE ISNULL([Event Name], '') <> ''
              ORDER BY [Event Name]";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        AddUnique(list, rdr["Event Name"].ToString());
                    }
                }

                string zoneSql =
                    @"SELECT [Name]
              FROM [Delivery Zones]
              WHERE ISNULL([Name], '') <> ''
              ORDER BY [Name]";

                using (SqlCommand cmd = new SqlCommand(zoneSql, conn))
                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        AddUnique(list, rdr["Name"].ToString());
                    }
                }
            }

            return Json(list.OrderBy(x => x).ToList());
        }

        private void AddUnique(List<string> list, string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return;

            if (!list.Any(x =>
                string.Equals(x, value, StringComparison.OrdinalIgnoreCase)))
            {
                list.Add(value);
            }
        }

        [HttpGet]
        public JsonResult GetDiscountAssignmentSelected(string eventName)
        {
            List<string> list = new List<string>();

            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql =
                    @"SELECT [Discount Name]
              FROM [Discount Events]
              WHERE [Event Name] = @EventName
              ORDER BY [POS]";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@EventName", eventName ?? "");

                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            list.Add(rdr["Discount Name"].ToString());
                        }
                    }
                }
            }

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveDiscountAssignmentEvent(
            [FromBody] DiscountAssignmentSaveRequest model)
        {
            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                using (SqlTransaction tran = conn.BeginTransaction())
                {
                    try
                    {
                        string deleteSql =
                            @"DELETE FROM [Discount Events]
                      WHERE [Event Name] = @EventName";

                        using (SqlCommand cmd =
                            new SqlCommand(deleteSql, conn, tran))
                        {
                            cmd.Parameters.AddWithValue("@EventName",
                                model.EventName ?? "");

                            cmd.ExecuteNonQuery();
                        }

                        string insertSql =
                            @"INSERT INTO [Discount Events]
                      ([Event Name], [POS], [Discount Name])
                      VALUES
                      (@EventName, @POS, @DiscountName)";

                        foreach (var item in model.Discounts)
                        {
                            using (SqlCommand cmd =
                                new SqlCommand(insertSql, conn, tran))
                            {
                                cmd.Parameters.AddWithValue("@EventName",
                                    model.EventName ?? "");

                                cmd.Parameters.AddWithValue("@POS",
                                    item.Pos);

                                cmd.Parameters.AddWithValue("@DiscountName",
                                    item.DiscountName ?? "");

                                cmd.ExecuteNonQuery();
                            }
                        }

                        tran.Commit();
                    }
                    catch
                    {
                        tran.Rollback();
                        throw;
                    }
                }
            }

            return Json(new { message = "Done" });
        }

        [HttpPost]
        public IActionResult DeleteDiscountAssignmentEvent(string eventName)
        {
            using (SqlConnection conn =
                new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                conn.Open();

                string sql1 =
                    @"DELETE FROM [Discount Events]
              WHERE [Event Name] = @EventName";

                using (SqlCommand cmd = new SqlCommand(sql1, conn))
                {
                    cmd.Parameters.AddWithValue("@EventName", eventName ?? "");
                    cmd.ExecuteNonQuery();
                }

                string sql2 =
                    @"DELETE FROM [Delivery Zones]
              WHERE [Name] = @EventName";

                using (SqlCommand cmd = new SqlCommand(sql2, conn))
                {
                    cmd.Parameters.AddWithValue("@EventName", eventName ?? "");
                    cmd.ExecuteNonQuery();
                }
            }

            return Json(new { message = "Done" });
        }
        // Paste inside your existing DiscountAdminController class.
        // Uses: using Microsoft.AspNetCore.Mvc; using System.Data; using System.Data.SqlClient; using Dapper;

        [HttpGet]
        public IActionResult Discounts()
        {
            return View("~/Views/MSaleAdmin/Discounts.cshtml");
        }

        private SqlConnection MConn() => new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

        [HttpGet]
        public async Task<IActionResult> GetDiscountSetupLists()
        {
            using var cn = MConn();
            var discounts = (await cn.QueryAsync<string>("SELECT [Discount Name] FROM [Discounts] ORDER BY [Discount Name]")).ToList();
            var departments = (await cn.QueryAsync<string>("SELECT [Report Name] FROM [Report Categories] ORDER BY [Report Name]")).ToList();
            var accessLevels = (await cn.QueryAsync<string>("SELECT [Access Name] FROM [Access Levels] ORDER BY [Access Name]")).ToList();
            return Json(new { discounts, departments, accessLevels });
        }

        [HttpGet]
        public async Task<IActionResult> GetDiscount(string discountName)
        {
            if (string.IsNullOrWhiteSpace(discountName)) return BadRequest("Missing discount name.");
            using var cn = MConn();

            var discount = await cn.QueryFirstOrDefaultAsync(@"
        SELECT [Discount Name] DiscountName, [Discount Type] DiscountType, [Discount Amount] DiscountAmount,
               [Access Level] AccessLevel, [Menu Item Link] MenuItemLink, [Category Link] CategoryLink,
               [Department Link] DepartmentLink, [Max Discount Amount] MaxDiscountAmount, [Match] MatchValue,
               [Taxable] Taxable, [Tax 1] Tax1, [Tax 2] Tax2, [Allow Only One] AllowOnlyOne,
               [After Print] AfterPrint, [Auto Match] AutoMatch, [Disc Department] DiscDepartment,
               [Type] Type, [Search Number] SearchNumber, [Start Time] StartTime, [End Time] EndTime,
               [Auto Deduct] AutoDeduct, [Expires] Expires, [Discount Id] DiscountId, [Tracking] Tracking
        FROM [Discounts] WHERE [Discount Name] = @discountName", new { discountName });

            var activeDays = await cn.QueryAsync(@"
        SELECT [POS] Pos, [Day] DayValue FROM [Active Days]
        WHERE [Discount Name] = @discountName", new { discountName });

            var assignments = await cn.QueryAsync<string>(@"
        SELECT [Link To Name] FROM [Coupon Assignments]
        WHERE [Coupon Name] = @discountName ORDER BY [Link To Name]", new { discountName });

            var comboItem = await cn.QueryAsync<string>(@"
        SELECT [Match] FROM [DiscountCombos]
        WHERE [Discount Name] = @discountName ORDER BY [Match]", new { discountName });

            var comboMatches = await cn.QueryAsync(@"
        SELECT [Match no] MatchNo, [Match] MatchName FROM [Match]
        WHERE [Discount Name] = @discountName ORDER BY [Match no], [Match]", new { discountName });

            return Json(new { discount, activeDays, assignments, comboItem, comboMatches });
        }

        [HttpGet]
        public async Task<IActionResult> GetDiscountAssignableItems(string mode)
        {
            using var cn = MConn();
            mode = (mode ?? "").ToLowerInvariant();
            var items = new List<string>();

            if (mode == "menu")
            {
                items = (await cn.QueryAsync<string>("SELECT [Menu Item Name] FROM [Menu Items] ORDER BY [Menu Item Name]")).ToList();
            }
            else if (mode == "category")
            {
                items.AddRange(await cn.QueryAsync<string>("SELECT [Category Name] FROM [Category Information] ORDER BY [Category Name]"));
                items.AddRange(await cn.QueryAsync<string>("SELECT [Sub Name] FROM [Sub Categories] ORDER BY [Sub Name]"));
                items = items.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct().OrderBy(x => x).ToList();
            }
            else if (mode == "department")
            {
                items = (await cn.QueryAsync<string>("SELECT [Report Name] FROM [Report Categories] ORDER BY [Report Name]")).ToList();
            }
            else if (mode == "combo")
            {
                items = (await cn.QueryAsync<string>("SELECT [Menu Item Name] FROM [Menu Items] ORDER BY [Menu Item Name]")).ToList();
                var mods = await cn.QueryAsync<string>("SELECT DISTINCT [Modifier Category] FROM [Modifiers] ORDER BY [Modifier Category]");
                items.AddRange(mods.Where(x => !string.IsNullOrWhiteSpace(x)));
                items = items.Distinct().OrderBy(x => x).ToList();
            }
            else return BadRequest("Mode must be menu, category, department, or combo.");

            return Json(items);
        }

        [HttpPost]
        public async Task<IActionResult> SaveDiscount([FromBody] DiscountSaveRequest r)
        {
            if (r == null || string.IsNullOrWhiteSpace(r.DiscountName)) return BadRequest("Missing discount name.");
            r.DiscountName = r.DiscountName.Replace("'", "").Trim();

            if ((r.DiscountType == "Charge Percentage" || r.DiscountType == "Charge Amount") && r.Type == "Discount")
                return BadRequest("Charge types cannot be saved as Type = Discount.");
            if (r.DiscountType == "Set Price" && (r.Assignments == null || r.Assignments.Count == 0))
                return BadRequest("Set Price requires assigned items.");

            using var cn = MConn();
            await cn.OpenAsync();
            using var tx = cn.BeginTransaction();
            try
            {
                var exists = await cn.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM [Discounts] WHERE [Discount Name] = @DiscountName", r, tx);
                if (exists == 0)
                {
                    await cn.ExecuteAsync(@"INSERT INTO [Discounts] ([Discount Name]) VALUES (@DiscountName)", r, tx);
                }

                await cn.ExecuteAsync(@"
            UPDATE [Discounts] SET
              [Discount Type]=@DiscountType, [Discount Amount]=@DiscountAmount, [Access Level]=@AccessLevel,
              [Menu Item Link]=@MenuItemLink, [Category Link]=@CategoryLink, [Department Link]=@DepartmentLink,
              [Max Discount Amount]=@MaxDiscountAmount, [Match]=@MatchValue, [Taxable]=@Taxable,
              [Tax 1]=@Tax1, [Tax 2]=@Tax2, [Allow Only One]=@AllowOnlyOne, [After Print]=@AfterPrint,
              [Auto Match]=@AutoMatch, [Disc Department]=@DiscDepartment, [Type]=@Type,
              [Search Number]=@SearchNumber, [Start Time]=@StartTime, [End Time]=@EndTime,
              [Auto Deduct]=@AutoDeduct, [Expires]=@Expires, [Discount Id]=@DiscountId, [Tracking]=@Tracking
            WHERE [Discount Name]=@DiscountName", r, tx);

                await cn.ExecuteAsync("DELETE FROM [Active Days] WHERE [Discount Name]=@DiscountName", r, tx);
                for (int i = 0; i < 7; i++)
                    await cn.ExecuteAsync("INSERT INTO [Active Days] ([Discount Name],[Day],[POS]) VALUES (@DiscountName,@DayValue,@Pos)",
                        new { r.DiscountName, DayValue = r.ActiveDays != null && r.ActiveDays.Contains(i) ? 1 : 0, Pos = i }, tx);

                await cn.ExecuteAsync("DELETE FROM [Coupon Assignments] WHERE [Coupon Name]=@DiscountName", r, tx);
                foreach (var item in (r.Assignments ?? new()).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct())
                    await cn.ExecuteAsync("INSERT INTO [Coupon Assignments] ([Coupon Name],[Link To Name]) VALUES (@DiscountName,@Item)", new { r.DiscountName, Item = item }, tx);

                await cn.ExecuteAsync("DELETE FROM [DiscountCombos] WHERE [Discount Name]=@DiscountName", r, tx);
                await cn.ExecuteAsync("DELETE FROM [Match] WHERE [Discount Name]=@DiscountName", r, tx);

                foreach (var item in (r.ComboItem ?? new()).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct())
                    await cn.ExecuteAsync(@"INSERT INTO [DiscountCombos] ([Discount Name],[Match],[Discount Amount],[Limit],[Coupon Type])
                                    VALUES (@DiscountName,@Item,@DiscountAmount,@MatchValue,@DiscountType)",
                        new { r.DiscountName, Item = item, r.DiscountAmount, r.MatchValue, r.DiscountType }, tx);

                foreach (var m in (r.ComboMatches ?? new()).Where(x => x.MatchNo >= 1 && x.MatchNo <= 6 && !string.IsNullOrWhiteSpace(x.MatchName)))
                    await cn.ExecuteAsync("INSERT INTO [Match] ([Match no],[Discount Name],[Match]) VALUES (@MatchNo,@DiscountName,@MatchName)",
                        new { m.MatchNo, r.DiscountName, m.MatchName }, tx);

                tx.Commit();
                return Json(new { ok = true });
            }
            catch (Exception ex)
            {
                tx.Rollback();
                return StatusCode(500, ex.Message);
            }
        }

        [HttpPost]
        public async Task<IActionResult> DeleteDiscount([FromBody] DiscountDeleteRequest r)
        {
            if (r == null || string.IsNullOrWhiteSpace(r.DiscountName)) return BadRequest("Missing discount name.");
            using var cn = MConn();
            await cn.OpenAsync();
            using var tx = cn.BeginTransaction();
            try
            {
                await cn.ExecuteAsync("DELETE FROM [Discounts] WHERE [Discount Name]=@DiscountName", r, tx);
                await cn.ExecuteAsync("DELETE FROM [DiscountCombos] WHERE [Discount Name]=@DiscountName", r, tx);
                await cn.ExecuteAsync("DELETE FROM [Match] WHERE [Discount Name]=@DiscountName", r, tx);
                await cn.ExecuteAsync("DELETE FROM [Coupon Assignments] WHERE [Coupon Name]=@DiscountName", r, tx);
                await cn.ExecuteAsync("DELETE FROM [Active Days] WHERE [Discount Name]=@DiscountName", r, tx);
                tx.Commit();
                return Json(new { ok = true });
            }
            catch (Exception ex)
            {
                tx.Rollback();
                return StatusCode(500, ex.Message);
            }
        }

        public sealed class DiscountDeleteRequest { public string DiscountName { get; set; } = ""; }
        public sealed class ComboMatchRequest { public int MatchNo { get; set; } public string MatchName { get; set; } = ""; }
        public sealed class DiscountSaveRequest
        {
            public string DiscountName { get; set; } = "";
            public string DiscountType { get; set; } = "";
            public decimal DiscountAmount { get; set; }
            public string AccessLevel { get; set; } = "";
            public int MenuItemLink { get; set; }
            public int CategoryLink { get; set; }
            public int DepartmentLink { get; set; }
            public decimal MaxDiscountAmount { get; set; }
            public decimal MatchValue { get; set; }
            public int Taxable { get; set; }
            public int Tax1 { get; set; } = 1;
            public int Tax2 { get; set; }
            public int AllowOnlyOne { get; set; }
            public int AfterPrint { get; set; }
            public int AutoMatch { get; set; }
            public string DiscDepartment { get; set; } = "None";
            public string Type { get; set; } = "Discount";
            public string SearchNumber { get; set; } = "None";
            public string StartTime { get; set; } = "12:00 AM";
            public string EndTime { get; set; } = "12:00 AM";
            public int AutoDeduct { get; set; }
            public string Expires { get; set; } = "";
            public int DiscountId { get; set; }
            public int Tracking { get; set; }
            public List<int> ActiveDays { get; set; } = new();
            public List<string> Assignments { get; set; } = new();
            public List<string> ComboItem { get; set; } = new();
            public List<ComboMatchRequest> ComboMatches { get; set; } = new();
        }


        // start of payroll summary code    
        

        public IActionResult PayrollSummary()
        {
            return View("~/Views/MSaleAdmin/PayrollSummary.cshtml");
        }
             
        
        // start of Time Record Editor
       
        public IActionResult TimeRecordEditor()
        {
            return View("~/Views/MSaleAdmin/TimeRecordEditor.cshtml");
        }
                
        // start of tip editor
        public IActionResult TipEditor()
        {
            return View("~/Views/MSaleAdmin/TipEditor.cshtml");
        }

        // start of time messages
        public IActionResult TimeMessages()
        {
            return View("~/Views/MSaleAdmin/TimeMessages.cshtml");
        }

        // Start of Check Review section
        public IActionResult CheckReview()
        {
            return View("~/Views/MSaleAdmin/CheckReview.cshtml");
        }              
       
        // start of modifier price levels code
        [HttpGet]
        public IActionResult ModifierPriceLevels()
        {
            return View();
        }           

        // start of modifier category maintenance code
        [HttpGet]
        public IActionResult ModifierDetails()
        {
            return View("~/Views/MSaleAdmin/ModifierDetails.cshtml");
        }
        // end of modifier category maintenance code
        
    }
}
