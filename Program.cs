using CAF.GstMatching.Business;
using CAF.GstMatching.Business.Interface;
using CAF.GstMatching.Helper;
using CAF.GstMatching;
using Microsoft.EntityFrameworkCore;
using CAF.GstMatching.Web.Common;
using System.Net;
using Microsoft.AspNetCore.Diagnostics;
using CAF.GstMatching.Web.SessionCheck;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews(options =>
{
    options.Filters.Add<SessionCheckFilter>(); // Register SessionCheckFilter globally
});

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddHttpClient();
builder.Services.AddSession(); // Session support
builder.Services.AddHttpContextAccessor(); // For MySession
// Register Business for Dependency Injection
builder.Services.AddScoped<IUserBusiness, UserBusiness>();

builder.Services.AddScoped<IPurchaseDataBusiness, PurchaseDataBusiness>();
builder.Services.AddScoped<IPurchaseTicketBusiness, PurchaseTicketBusiness>();
builder.Services.AddScoped<IGSTR2DataBusiness, GSTR2DataBusiness>();
builder.Services.AddScoped<ICompareGstBusiness, CompareGstBusiness>();
builder.Services.AddScoped<IModifiedDataBusiness, ModifiedDataBusiness>();

builder.Services.AddScoped<ISLDataBusiness,SLDataBusiness>();
builder.Services.AddScoped<ISLEInvoiceBusiness, SLEInvoiceBusiness>();
builder.Services.AddScoped<ISLEWayBillBusiness, SLEWayBillBusiness>();
builder.Services.AddScoped<ISLTicketsBusiness, SLTicketsBusiness>();
builder.Services.AddScoped<ISLComparedDataBusiness, SLComparedDataBusiness>();

builder.Services.AddScoped<INoticeDataBusiness, NoticeDataBusiness>();

builder.Services.AddScoped<UserHelper>();

builder.Services.AddScoped<PurchaseDataHelper>(); 
builder.Services.AddScoped<PurchaseTicketHelper>();
builder.Services.AddScoped<GSTR2DataHelper>();
builder.Services.AddScoped<CompareGstHelper>();
builder.Services.AddScoped<ModifiedDataHelper>();

builder.Services.AddScoped<SLDataHelper>();
builder.Services.AddScoped<SLEInvoiceHelper>();
builder.Services.AddScoped<SLEWayBillHelper>();
builder.Services.AddScoped<SLTicketsHelper>();
builder.Services.AddScoped<SLComparedDataHelper>();

builder.Services.AddScoped<NoticeDataHelper>();

builder.Services.AddScoped<ApplicationDbContext>();

builder.Services.AddDbContext<CAF.GstMatching.Web.ApplicationDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

builder.Services.AddDistributedMemoryCache();
builder.Services.AddSession(options =>
{
    options.IdleTimeout = TimeSpan.FromMinutes(30); // Session timeout duration
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
});

builder.Services.AddSingleton<IHttpContextAccessor, HttpContextAccessor>();

// Add IHttpContextAccessor
builder.Services.AddHttpContextAccessor();
builder.Logging.AddConsole(); // Ensure console logging
builder.Logging.AddDebug();

builder.Services.AddSignalR(); // ✅ ADD THIS SignalR SERVICE

var app = builder.Build();

// Configure the HTTP request pipeline.
//if (app.Environment.IsDevelopment())
//{
//    app.UseDeveloperExceptionPage();
//}
//else
//{
//    app.UseExceptionHandler("/Home/Error");
//    app.UseHsts();
//}

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler(errorApp =>
    {
        errorApp.Run(async context =>
        {
            context.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
            context.Response.ContentType = "text/html";
            var exceptionHandlerPathFeature = context.Features.Get<IExceptionHandlerPathFeature>();
            if (exceptionHandlerPathFeature?.Error != null)
            {
                var errorMessage = exceptionHandlerPathFeature.Error.Message;
                await context.Response.WriteAsync($"<html><body>\n<h1>Error occurred</h1>\n<pre>{errorMessage}</pre>\n</body></html>");
            }
        });
    });
    app.UseHsts();
}

MySession.Configure(app.Services.GetRequiredService<IHttpContextAccessor>());

app.UseRouting();

app.Use(async (context, next) =>
{
    context.Response.Headers["Cache-Control"] = "no-cache, no-store, must-revalidate";
    context.Response.Headers["Pragma"] = "no-cache";
    context.Response.Headers["Expires"] = "0";
    await next();
});

app.UseHttpsRedirection();
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        ctx.Context.Response.Headers["Cache-Control"] = "no-cache, no-store, must-revalidate, private"; // Added cache headers for static files
        ctx.Context.Response.Headers["Pragma"] = "no-cache";
        ctx.Context.Response.Headers["Expires"] = "0";
    }
});
app.UseSession();
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.MapHub<CAF.GstMatching.Web.Hubs.ChatHub>("/chathub"); // ✅ ADD THIS


app.Run();
