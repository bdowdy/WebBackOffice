using Microsoft.AspNetCore.Mvc;

namespace MSaleWebServer.Controllers
{
    public partial class MSaleAdminController : Controller
    {
        // Serves the Language / Translation editor page.
        // Data is loaded and saved client-side via /api/language.
        public IActionResult LanguageEditor()
        {
            return View();
        }
    }
}
