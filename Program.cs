using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using MSaleWebServer.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllersWithViews();
builder.Services.AddSingleton<SmsService>();

// API documentation (Swagger / OpenAPI)
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Swagger UI at /swagger
app.UseSwagger();
app.UseSwaggerUI();

app.UseStaticFiles();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=MSaleAdmin}/{action=Index}/{id?}");

// Keep API endpoints working too
app.MapControllers();

app.Run();