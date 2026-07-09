using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace MSaleWebServer.Controllers
{
    [AllowAnonymous]
    public class AccountController : Controller
    {
        private readonly IConfiguration _config;
        public AccountController(IConfiguration config) => _config = config;

        [HttpGet]
        public IActionResult Login(string? returnUrl = null)
        {
            ViewBag.ReturnUrl = returnUrl;
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(string username, string password, string? returnUrl = null)
        {
            var cfgUser = _config["Admin:Username"] ?? "";
            var cfgPass = _config["Admin:Password"] ?? "";

            if (cfgUser.Length == 0 || cfgPass.Length == 0)
            {
                ViewBag.Error = "No admin credentials are configured on the server.";
                return View();
            }

            if (!FixedTimeEquals(username, cfgUser) || !FixedTimeEquals(password, cfgPass))
            {
                ViewBag.Error = "Invalid username or password.";
                ViewBag.ReturnUrl = returnUrl;
                return View();
            }

            var identity = new ClaimsIdentity(
                new[] { new Claim(ClaimTypes.Name, username) },
                CookieAuthenticationDefaults.AuthenticationScheme);

            await HttpContext.SignInAsync(
                CookieAuthenticationDefaults.AuthenticationScheme,
                new ClaimsPrincipal(identity),
                new AuthenticationProperties { IsPersistent = true });

            return Redirect(string.IsNullOrEmpty(returnUrl) || !Url.IsLocalUrl(returnUrl)
                ? "/MSaleAdmin" : returnUrl);
        }

        [HttpPost]
        public async Task<IActionResult> Logout()
        {
            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
            return RedirectToAction("Login");
        }

        private static bool FixedTimeEquals(string a, string b) =>
            CryptographicOperations.FixedTimeEquals(
                SHA256.HashData(Encoding.UTF8.GetBytes(a)),
                SHA256.HashData(Encoding.UTF8.GetBytes(b)));
    }
}
