using Microsoft.AspNetCore.Mvc;

namespace MSaleWebServer.Controllers
{
    public partial class MSaleAdminController : Controller
    {
        // Serves the Receipt Print Size editor (large/small for menu items
        // and modifiers). Data is loaded/saved via /api/receipt-esc.
        public IActionResult ReceiptPrintSize()
        {
            return View();
        }
    }
}
