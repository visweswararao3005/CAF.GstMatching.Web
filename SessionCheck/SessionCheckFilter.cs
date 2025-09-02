using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;

namespace CAF.GstMatching.Web.SessionCheck
{
    public class SessionCheckFilter : IActionFilter
    {
        public void OnActionExecuting(ActionExecutingContext context)
        {
            // Log request details
            var action = context.ActionDescriptor.RouteValues["action"]?.ToLower();
            var controller = context.ActionDescriptor.RouteValues["controller"]?.ToLower();
            //Console.WriteLine($"SessionCheckFilter: Controller={controller}, Action={action}, Path={context.HttpContext.Request.Path}");

            // Skip session check for public actions
            if (controller == "home" && (action == "index" || action == "login" || action == "signup" || action == "checksession"))
            {
                //Console.WriteLine("Skipping session check for public page");
                return;
            }

            // Check session validity
            if (!context.HttpContext.Session.TryGetValue("UserId", out _) || !context.HttpContext.Session.Keys.Any())
            {
                var basePath = context.HttpContext.Request.PathBase.ToString();
                var loginPath = $"{basePath}/Home/Login".TrimEnd('/');
                //Console.WriteLine($"Redirecting to: {loginPath}");
                context.Result = new RedirectResult(loginPath);
            }
        }

        public void OnActionExecuted(ActionExecutedContext context)
        {
            //Console.WriteLine("SessionCheckFilter: Action executed");
        }
    }
}