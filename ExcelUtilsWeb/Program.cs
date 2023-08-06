using ExcelUtilsWeb.Services;
using ExcelUtilsWeb.Services.Interfaces;
using System.Diagnostics;

var builder = WebApplication.CreateBuilder(args);
var useUrl = builder.Configuration["Url"];
builder.WebHost.UseUrls(useUrl);

// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddAntDesign();

builder.Services.AddScoped<IExcelService, ExcelService>()
    .AddScoped<IFileService, FileService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
}


app.UseStaticFiles();

app.UseRouting();

app.MapBlazorHub();
app.MapFallbackToPage("/_Host");

app.Lifetime.ApplicationStarted.Register(() =>
{
    OpenBrowser(useUrl);
});

app.Run();

static void OpenBrowser(string url)
{
    Process.Start(
        new ProcessStartInfo("cmd", $"/c start {url}")
        {
            CreateNoWindow = true
        });
}
