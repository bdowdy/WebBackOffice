using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    public class PrinterSetupController : Controller
    {
        private readonly IConfiguration _configuration;

        public PrinterSetupController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View("~/Views/MSaleAdmin/PrinterSetup.cshtml");
        }

        [HttpGet]
        public IActionResult GetSystemPrinters()
        {
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            var list = cn.Query<SystemPrinterDto>(@"
                SELECT
                    [RegisterName]   AS RegisterName,
                    [Printer Name]   AS PrinterName,
                    [Port Name]      AS PortName,
                    [Data Rate]      AS DataRate,
                    [Failed]         AS Failed,
                    [Go Order]       AS GoOrder,
                    [Delivery]       AS Delivery,
                    [Fire Order]     AS FireOrder,
                    [Printer Group]  AS PrinterGroup
                FROM [System Printers]
                ORDER BY [Printer Name];
            ").ToList();

            return Json(list);
        }

        [HttpPost]
        public IActionResult SaveSystemPrinter([FromBody] SystemPrinterDto x)
        {
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Execute(@"
                IF EXISTS (SELECT 1 FROM [System Printers] WHERE [Printer Name] = @PrinterName)
                BEGIN
                    UPDATE [System Printers]
                    SET
                        [Port Name]   = @PortName,
                        [Data Rate]   = @DataRate,
                        [Failed]      = @Failed,
                        [Go Order]    = @GoOrder,
                        [Delivery]    = @Delivery,
                        [Fire Order]  = @FireOrder,
                        [Printer Group] = @PrinterGroup
                    WHERE [Printer Name] = @PrinterName;
                END
                ELSE
                BEGIN
                    INSERT INTO [System Printers]
                    (
                        [Printer Name],
                        [Port Name],
                        [Data Rate],
                        [Failed],
                        [Go Order],
                        [Delivery],
                        [Fire Order],
                        [Printer Group]
                    )
                    VALUES
                    (
                        @PrinterName,
                        @PortName,
                        @DataRate,
                        @Failed,
                        @GoOrder,
                        @Delivery,
                        @FireOrder,
                        @PrinterGroup
                    );
                END
            ", x);

            return Json(new { ok = true });
        }

        [HttpPost]
        public IActionResult DeleteSystemPrinter([FromBody] SystemPrinterDto x)
        {
            using var cn = new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

            cn.Execute(@"
                DELETE FROM [System Printers]
                WHERE [Printer Name] = @PrinterName;
            ", x);

            return Json(new { ok = true });
        }
    }

    public class SystemPrinterDto
    {
        public string RegisterName  { get; set; } = "";
        public string PrinterName   { get; set; } = "";
        public string PortName      { get; set; } = "";
        public string DataRate      { get; set; } = "";
        public string Failed        { get; set; } = "";
        public string GoOrder       { get; set; } = "";
        public string Delivery      { get; set; } = "";
        public bool   FireOrder     { get; set; }
        public string PrinterGroup  { get; set; } = "";
    }
}
