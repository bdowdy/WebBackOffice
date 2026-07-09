using System.Net.Http.Json;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using MSaleWebServer.Services;

var builder = WebApplication.CreateBuilder(args);

// Allows running as a Windows Service on the cloud server (no-op when run normally)
builder.Host.UseWindowsService();

builder.Services.AddControllersWithViews();
builder.Services.AddSingleton<SmsService>();
builder.Services.AddSingleton<TerminalUpdateService>();

// ── Authentication ────────────────────────────────────────────────────────────
// Back-office pages: cookie login (credentials in user secrets / env vars).
// Sync API: X-Api-Key header (see middleware below).
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(o =>
    {
        o.LoginPath = "/Account/Login";
        o.ExpireTimeSpan = TimeSpan.FromHours(10);
        o.SlidingExpiration = true;
    });

// Everything requires a signed-in user unless marked [AllowAnonymous]
builder.Services.AddAuthorization(o =>
    o.FallbackPolicy = new AuthorizationPolicyBuilder()
        .RequireAuthenticatedUser()
        .Build());

// API documentation (Swagger / OpenAPI)
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Swagger only in Development - not on the public server
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseStaticFiles();

app.UseAuthentication();

// ── API key auth for the sync agent ──────────────────────────────────────────
// A request with a valid X-Api-Key header is treated as the authenticated
// sync agent. Key lives in user secrets / env var Sync__ApiKey.
app.Use(async (ctx, next) =>
{
    if (ctx.User.Identity?.IsAuthenticated != true &&
        ctx.Request.Headers.TryGetValue("X-Api-Key", out var supplied))
    {
        var expected = app.Configuration["Sync:ApiKey"];
        if (!string.IsNullOrWhiteSpace(expected) &&
            CryptographicOperations.FixedTimeEquals(
                SHA256.HashData(Encoding.UTF8.GetBytes(supplied.ToString())),
                SHA256.HashData(Encoding.UTF8.GetBytes(expected))))
        {
            ctx.User = new ClaimsPrincipal(new ClaimsIdentity(
                new[] { new Claim(ClaimTypes.Name, "sync-agent") }, "ApiKey"));
        }
    }
    await next();
});

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=MSaleAdmin}/{action=Index}/{id?}");

// Keep API endpoints working too
app.MapControllers();

// ── Startup: open the hosted sales dashboard, signed in ──────────────────────
// Configure in appsettings.json under "Dashboard". Disable: OpenOnStart=false.
if (app.Configuration.GetValue("Dashboard:OpenOnStart", true))
{
    app.Lifetime.ApplicationStarted.Register(() => _ = Task.Run(async () =>
    {
        var cfg        = app.Configuration;
        var url        = cfg["Dashboard:SalesUrl"] ?? "https://app.microsaledashboard.net/dashboard";
        var tokenParam = cfg["Dashboard:TokenParam"] ?? "token";
        var token      = cfg["Dashboard:ApiKey"];   // static key option

        // Or exchange username/password for a short-lived token server-to-server
        var user     = cfg["Dashboard:Username"];
        var pass     = cfg["Dashboard:Password"];
        var loginUrl = cfg["Dashboard:LoginUrl"];
        if (string.IsNullOrWhiteSpace(token) &&
            !string.IsNullOrWhiteSpace(user) && !string.IsNullOrWhiteSpace(loginUrl))
        {
            try
            {
                using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
                var resp = await http.PostAsJsonAsync(loginUrl, new { username = user, password = pass });
                if (resp.IsSuccessStatusCode)
                {
                    var doc = await resp.Content.ReadFromJsonAsync<System.Text.Json.JsonDocument>();
                    var root = doc!.RootElement;
                    foreach (var field in new[] { "token", "accessToken", "access_token" })
                        if (root.TryGetProperty(field, out var t)) { token = t.GetString(); break; }
                }
                else
                    app.Logger.LogWarning("Dashboard auto-login failed: {Status}", resp.StatusCode);
            }
            catch (Exception ex)
            {
                app.Logger.LogWarning(ex, "Dashboard auto-login failed - opening without sign-in.");
            }
        }

        if (!string.IsNullOrWhiteSpace(token))
            url += (url.Contains('?') ? "&" : "?") + $"{tokenParam}={Uri.EscapeDataString(token)}";

        // Open two tabs: back office first, sales dashboard second (so it's focused)
        var backOffice = cfg["Dashboard:BackOfficeUrl"];
        if (string.IsNullOrWhiteSpace(backOffice))
            backOffice = app.Urls.FirstOrDefault(u => u.StartsWith("http://"))
                         ?? app.Urls.FirstOrDefault()
                         ?? "http://localhost:5088";
        app.Logger.LogInformation("Opening back office: {BackOffice}", backOffice);
        app.Logger.LogInformation("Opening sales dashboard: {Sales}", url);
        try
        {
            System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo(backOffice) { UseShellExecute = true });
            await Task.Delay(1500);   // let the first tab open so order is stable
            System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch { /* headless/service session - nothing to open */ }
    }));
}

app.Run();