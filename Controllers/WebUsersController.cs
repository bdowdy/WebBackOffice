using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MSaleWebServer.Services;

namespace MSaleWebServer.Controllers
{
    /// <summary>
    /// Manage web back-office logins (dbo.WebUsers).
    /// Page: /WebUsers   APIs: /WebUsers/List, /WebUsers/Save, /WebUsers/Delete
    /// Protected by the global cookie-auth fallback policy.
    /// </summary>
    public class WebUsersController : Controller
    {
        private readonly IConfiguration _configuration;

        public WebUsersController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View("~/Views/MSaleAdmin/WebUsers.cshtml");
        }

        [HttpGet]
        public JsonResult List()
        {
            var list = new List<object>();

            using (SqlConnection conn = Open())
            {
                string sql = @"SELECT Username, IsActive, CreatedAt
                               FROM dbo.WebUsers
                               ORDER BY Username";

                using SqlCommand cmd = new SqlCommand(sql, conn);
                using SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    list.Add(new
                    {
                        username = rdr["Username"].ToString(),
                        isActive = (bool)rdr["IsActive"],
                        createdAt = ((DateTime)rdr["CreatedAt"]).ToString("yyyy-MM-dd")
                    });
                }
            }

            return Json(list);
        }

        public class SaveWebUserModel
        {
            public string OldUsername { get; set; } = "";
            public string Username { get; set; } = "";
            public string Password { get; set; } = "";
            public bool IsActive { get; set; } = true;
        }

        [HttpPost]
        public IActionResult Save([FromBody] SaveWebUserModel model)
        {
            if (string.IsNullOrWhiteSpace(model.Username))
                return Content("Username is required.");

            bool isNew = string.IsNullOrEmpty(model.OldUsername);

            if (isNew && string.IsNullOrWhiteSpace(model.Password))
                return Content("Password is required for a new user.");

            // Don't let someone deactivate the account they are signed in with
            if (!isNew &&
                string.Equals(model.OldUsername, User.Identity?.Name, StringComparison.OrdinalIgnoreCase) &&
                !model.IsActive)
                return Content("You cannot deactivate the account you are signed in with.");

            using (SqlConnection conn = Open())
            {
                if (isNew)
                {
                    using SqlCommand exists = new SqlCommand(
                        "SELECT COUNT(*) FROM dbo.WebUsers WHERE Username = @U", conn);
                    exists.Parameters.AddWithValue("@U", model.Username);
                    if ((int)exists.ExecuteScalar() > 0)
                        return Content("That username already exists.");

                    string insertSql = @"INSERT INTO dbo.WebUsers (Username, PasswordHash, IsActive)
                                         VALUES (@Username, @Hash, @IsActive)";
                    using SqlCommand cmd = new SqlCommand(insertSql, conn);
                    cmd.Parameters.AddWithValue("@Username", model.Username);
                    cmd.Parameters.AddWithValue("@Hash", WebUserStore.HashPassword(model.Username, model.Password));
                    cmd.Parameters.AddWithValue("@IsActive", model.IsActive);
                    cmd.ExecuteNonQuery();
                }
                else
                {
                    string updateSql = string.IsNullOrWhiteSpace(model.Password)
                        ? @"UPDATE dbo.WebUsers
                            SET Username = @Username, IsActive = @IsActive
                            WHERE Username = @OldUsername"
                        : @"UPDATE dbo.WebUsers
                            SET Username = @Username, IsActive = @IsActive, PasswordHash = @Hash
                            WHERE Username = @OldUsername";

                    using SqlCommand cmd = new SqlCommand(updateSql, conn);
                    cmd.Parameters.AddWithValue("@Username", model.Username);
                    cmd.Parameters.AddWithValue("@IsActive", model.IsActive);
                    cmd.Parameters.AddWithValue("@OldUsername", model.OldUsername);
                    if (!string.IsNullOrWhiteSpace(model.Password))
                        cmd.Parameters.AddWithValue("@Hash", WebUserStore.HashPassword(model.Username, model.Password));

                    if (cmd.ExecuteNonQuery() == 0)
                        return Content("User not found.");
                }
            }

            return Content("Saved Successfully");
        }

        public class DeleteWebUserModel
        {
            public string Username { get; set; } = "";
        }

        [HttpPost]
        public IActionResult Delete([FromBody] DeleteWebUserModel model)
        {
            if (string.Equals(model.Username, User.Identity?.Name, StringComparison.OrdinalIgnoreCase))
                return Content("You cannot delete the account you are signed in with.");

            using (SqlConnection conn = Open())
            {
                using SqlCommand cmd = new SqlCommand(
                    "DELETE FROM dbo.WebUsers WHERE Username = @U", conn);
                cmd.Parameters.AddWithValue("@U", model.Username);
                cmd.ExecuteNonQuery();
            }

            return Content("Deleted Successfully");
        }

        private SqlConnection Open()
        {
            var conn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
            conn.Open();
            WebUserStore.EnsureTable(conn);
            return conn;
        }
    }
}
