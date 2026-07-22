using Microsoft.AspNetCore.Identity;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Services
{
    /// <summary>
    /// Database-backed back-office logins (dbo.WebUsers).
    /// Passwords are stored as ASP.NET Identity PBKDF2 hashes - never plain text.
    /// Config-based logins (Admin:Username / Admin:Users) keep working as a fallback,
    /// so a database problem can never lock everyone out.
    /// </summary>
    public static class WebUserStore
    {
        private static readonly PasswordHasher<string> Hasher = new();

        public static void EnsureTable(SqlConnection conn)
        {
            const string sql = @"
IF OBJECT_ID('dbo.WebUsers', 'U') IS NULL
CREATE TABLE dbo.WebUsers (
    Id           INT IDENTITY(1,1) PRIMARY KEY,
    Username     NVARCHAR(100)  NOT NULL UNIQUE,
    PasswordHash NVARCHAR(500)  NOT NULL,
    IsActive     BIT            NOT NULL DEFAULT 1,
    CreatedAt    DATETIME2      NOT NULL DEFAULT SYSUTCDATETIME()
)";
            using var cmd = new SqlCommand(sql, conn);
            cmd.ExecuteNonQuery();
        }

        public static string HashPassword(string username, string password) =>
            Hasher.HashPassword(username, password);

        /// <summary>True when the username/password matches an active database user.</summary>
        public static bool TryValidate(IConfiguration config, string username, string password)
        {
            try
            {
                var cs = config.GetConnectionString("DefaultConnection");
                if (string.IsNullOrWhiteSpace(cs)) return false;

                using var conn = new SqlConnection(cs);
                conn.Open();
                EnsureTable(conn);

                using var cmd = new SqlCommand(
                    "SELECT PasswordHash FROM dbo.WebUsers WHERE Username = @U AND IsActive = 1", conn);
                cmd.Parameters.AddWithValue("@U", username ?? "");

                if (cmd.ExecuteScalar() is not string hash || hash.Length == 0) return false;

                return Hasher.VerifyHashedPassword(username, hash, password ?? "")
                       != PasswordVerificationResult.Failed;
            }
            catch
            {
                // Database unavailable - fall back to config logins only
                return false;
            }
        }
    }
}
