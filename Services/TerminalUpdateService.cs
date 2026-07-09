using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Services
{
    // Flags POS terminals to reload after a web back-office change.
    //
    // Each terminal, on logon, checks [Terminal Update] for a row where
    // [Name] = MyRegisterName and [date] <= now (see TerminalUpdates.bas).
    // So after any change here, we insert one such row per terminal.
    //
    // [Terminal Update].[Name] must equal MyRegisterName. In the POS,
    // MyRegisterName is matched against [register locations].[Register Name],
    // so that is the authoritative list of names to insert.
    public class TerminalUpdateService
    {
        private readonly IConfiguration _configuration;
        public TerminalUpdateService(IConfiguration configuration) => _configuration = configuration;

        // Source of register names to flag. Default = [register locations].
        // Alternative (if your MachineName equals the register name):
        //   SELECT DISTINCT [MachineName] AS Name FROM [dbo].[Terminals] ...
        private const string RegisterNameSource =
            "SELECT DISTINCT [Register Name] AS Name FROM [dbo].[register locations] " +
            "WHERE [Register Name] IS NOT NULL AND LTRIM(RTRIM([Register Name])) <> ''";

        /// <summary>
        /// Insert a pending update row (dated now) for every terminal that
        /// does not already have one. Never throws - flagging must not break
        /// the save that triggered it. Returns the number of terminals flagged.
        /// </summary>
        public int FlagAllTerminals()
        {
            try
            {
                using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));
                cn.Open();

                // One pending row per terminal (skip terminals already flagged),
                // so rows don't pile up before a terminal next logs on.
                string sql =
                    "INSERT INTO [dbo].[Terminal Update] ([Name], [date]) " +
                    "SELECT src.Name, GETDATE() FROM (" + RegisterNameSource + ") AS src " +
                    "WHERE NOT EXISTS (SELECT 1 FROM [dbo].[Terminal Update] tu WHERE tu.[Name] = src.Name);";

                using var cmd = new SqlCommand(sql, cn);
                return cmd.ExecuteNonQuery();
            }
            catch
            {
                // Swallow: the change already saved successfully; a failed flag
                // should not surface as a save error. (Add logging if desired.)
                return 0;
            }
        }
    }
}
