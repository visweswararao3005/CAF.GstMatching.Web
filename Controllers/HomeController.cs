using CAF.GstMatching.Business.Interface;
using CAF.GstMatching.Web.Common;
using CAF.GstMatching.Web.Helpers;
using CAF.GstMatching.Web.Models;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Net.Http;
using System.Security.Claims;
using System.Text;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;
using System.Net.Mail;
using CAF.GstMatching.Models.UserModel;
using System.Net;

namespace CAF.GstMatching.Web.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IHttpClientFactory _httpClientFactory;
        //private readonly HttpClient _httpClient;
        private readonly IConfiguration _configuration;
        private readonly IPurchaseTicketBusiness _purchaseTicketBusiness;
        private readonly IUserBusiness _userBusiness;

        public HomeController(ILogger<HomeController> logger,
                              IPurchaseTicketBusiness purchaseTicketBusiness,
                              IHttpClientFactory httpClientFactory,
                              IUserBusiness userBusiness,
                              IConfiguration configuration)
        {
            _logger = logger;
            //_httpClient = new HttpClient();
            _purchaseTicketBusiness = purchaseTicketBusiness;
            _httpClientFactory = httpClientFactory;
            _configuration = configuration;
            _userBusiness = userBusiness;
        }

        #region Index
        public async Task<IActionResult> Index()
        {
            await TjCaptions("Homepage");
            return View();
        }

        #endregion

        #region KeepAlive
        [HttpGet]
		public IActionResult KeepAlive()
		{
			return Ok();
		}

        #endregion     

        #region Error
        public async Task<IActionResult> Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        #endregion

        #region Login
        // GET: /Home/Login (Page Load)
        public async Task<IActionResult> Login()
        {
			//var emails = await _userBusiness.GetAdminUsers(); // Get email list for validation
			//Console.WriteLine("Email List: " + string.Join(", ", emails.Select(e => e.Email)));
			if (TempData["Messages"] != null)
            {
                ViewBag.Message = TempData["Messages"]; // From Signup
            }
            await TjCaptions("Loginpage");
            return View();
        }

        // POST: /Home/Login (Form Submit)
        [HttpPost]
        public async Task<IActionResult> Login(string email, string password, string userType)
        {
            var emails = await _userBusiness.GetAdminUsers(); // Get email list for validation
            var mainAdminEmails = await _userBusiness.GetMainAdminUsers(); // Get email list for validation
			//Console.WriteLine("Email List: " + string.Join(", ", emails.Select(e => e.Email)));

			// Store form values to repopulate the form in case of an error
			//var formValues = new Dictionary<string, string>
   //         {
   //             { "Email", email ?? "" },
   //             { "password", password ?? "" }
   //         };
   //         ViewBag.FormValues = formValues;

            // Load captions for the login page
            await TjCaptions("Loginpage");

            if (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(password))
            {
                ViewBag.Message = "Email and Password are required";
                // Store form values to repopulate the form in case of an error
                var formValues = new Dictionary<string, string>
                {
                    { "Email", email ?? "" },
                    { "password", password ?? "" }
                };
                ViewBag.FormValues = formValues;
                return View("Login");
            }

            var requestBody = new { LanguageId = "EN", EmailId = email, Password = password };
            var content = new StringContent(JsonConvert.SerializeObject(requestBody), Encoding.UTF8, "application/json");
            try
            {
                string baseUrl = _configuration["ApiSettings:BaseUrl"] ?? throw new InvalidOperationException("BaseUrl not configured");
                string validateLoginEndpoint = _configuration["ApiSettings:ValidateLoginEndpoint"] ?? throw new InvalidOperationException("ValidateLoginEndpoint not configured");
                string apiUrl = $"{baseUrl}{validateLoginEndpoint}";

                var httpClient = _httpClientFactory.CreateClient();
                var response = await httpClient.PostAsync(apiUrl, content);
                var responseData = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    var user = JsonConvert.DeserializeObject<User>(responseData);
                    if (user == null || string.IsNullOrEmpty(user.userName))
                    {
                        ViewBag.Message = "Login failed: Invalid response from server";
                        // Store form values to repopulate the form in case of an error
                        var formValues = new Dictionary<string, string>
                        {
                            { "Email", email ?? "" },
                            { "password", password ?? "" }
                        };
                        ViewBag.FormValues = formValues;
                        return View("Login");
                    }

                    //_logger.LogInformation("User : name-{0} ,mail-{1} ,Code-{2}, gstin-{3}, password-{4}",
                        //user.userName, user.emailId, user.userCode, user.gstin, password);
                    //_logger.LogInformation("password from form{0}", password);
                    MySession.Current.UserName = user.firstName;
                    MySession.Current.Email = user.emailId;
                    MySession.Current.UserCode = user.userCode;
                    MySession.Current.gstin = user.gstin;
                    MySession.Current.loginpassword = password;

                    //_logger.LogInformation("gstin from session{0}", user.gstin);
                    //_logger.LogInformation("Name from session{0}", user.userName);
                    // Set UserId in session for data-logged-in check
                    HttpContext.Session.SetString("UserId", user.userCode);

                    await TjCaptions("Dashboardpage"); // Load dashboard captions

                    // Return different views based on userType
                    var userValidUpto = await _userBusiness.getUserValidUpto(email);
                    bool valid = (userValidUpto.accessToDate != null &&  DateTime.Now <= userValidUpto.accessToDate);
                    if (emails.Any(e => e.Email.Equals(email, StringComparison.OrdinalIgnoreCase)))
					{
                        if (valid)
                        {
                            if (user.passwordChanged == "No")
                            {
                                ViewBag.Messages = "Admin";
                                ViewBag.UserName = user.userName;
                                return RedirectToAction("ChangePassword", "Admin");
                            }
                            ViewBag.Messages = "Admin";
                            ViewBag.UserName = user.userName;
                            return RedirectToAction("OpenTaskCSV", "Admin");
                        }
                        else
                        {
                            return RedirectToAction("Index", "Payment");
                        }
                    }
                    else if (mainAdminEmails.Any(e => e.Email.Equals(email, StringComparison.OrdinalIgnoreCase)))
                    {
                        if (user.passwordChanged == "No")
                        {
                            ViewBag.Messages = "MainAdmin";
                            ViewBag.UserName = user.userName;
                            return RedirectToAction("ChangePassword", "MainAdmin");
                        }
                        ViewBag.Messages = "MainAdmin";
                        ViewBag.UserName = user.userName;
                        return RedirectToAction("ActiveNotices", "MainAdmin");
                    }
                    else
                    {
                        if (valid)
                        {
                            if (user.passwordChanged == "No")
                            {
                                ViewBag.Messages = "Vendor";
                                ViewBag.UserName = user.userName;
                                return RedirectToAction("ChangePasswordV", "Vendor");
                            }
                            ViewBag.Messages = "Vendor";
                            ViewBag.UserName = user.userName;
                            return RedirectToAction("DashboardView", "Vendor"); // Vendor dashboard
                        }
                        else
                        {
                            return RedirectToAction("Index", "Payment");
                        }
                    }
                }
                else
                {
                    ViewBag.Message = "Invalid Username or Password";
                    // Store form values to repopulate the form in case of an error
                    var formValues = new Dictionary<string, string>
                    {
                        { "Email", email ?? "" },
                        { "password", password ?? "" }
                    };
                    ViewBag.FormValues = formValues;
                    return View("Login");
                }
            }
            catch (Exception)
            {
                //_logger.LogError(ex, "Error during ValidateLogin");
                ViewBag.Message = "Something went wrong";
                // Store form values to repopulate the form in case of an error
                var formValues = new Dictionary<string, string>
                {
                    { "Email", email ?? "" },
                    { "password", password ?? "" }
                };
                ViewBag.FormValues = formValues;
                return View("Login");
            }
        }

        #endregion

        #region LogOut
        //public void LogOut()
        //{
        //    Response.Redirect("/techiejoeweb/?mode=logout");
        //    HttpContext.Session.Clear();
        //    ViewData["ActiveAction"] = "LogOut";
        //}

        #endregion

        #region Signup
        public async Task<IActionResult> Signup(string lblBussinesseMail, string lblFullName, string lblOrganizationName, string ddlStateCode, string txtPAN, string txtEntity, string txtZ, string txtChecksum, string lblGSTIN, string lblAddress, string lblPhoneNo)
        {
            var stateCodeList = await _userBusiness.StateCodeList(); // ? Await the async method
            ViewBag.StateCodeList = stateCodeList;
            // Store form values to repopulate the form in case of an error or success
            var formValues = new Dictionary<string, string>
            {
                { "lblBussinesseMail", lblBussinesseMail },
                { "lblFullName", lblFullName },
                { "lblOrganizationName", lblOrganizationName },
                { "ddlStateCode", ddlStateCode },
                { "txtPAN", txtPAN },
                { "txtEntity", txtEntity },
                { "txtZ", txtZ },
                { "txtChecksum", txtChecksum },
                { "lblGSTIN", lblGSTIN },
                { "lblAddress", lblAddress },
                { "lblPhoneNo", lblPhoneNo }
            };
            ViewBag.FormValues = formValues;

            // Load captions for the signup page initially
            await TjCaptions("Signuppage");

            if (lblBussinesseMail != null && lblFullName != null)
            {
                var httpClient = _httpClientFactory.CreateClient();
                var requestBody = new
                {
                    LanguageId = "EN",
                    BusinessEmailAddress = lblBussinesseMail,
                    FullName = lblFullName,
                    OrganizationName = lblOrganizationName,
                    Position = lblGSTIN,
                    Address = lblAddress,
                    PhoneNumber = lblPhoneNo,
                    password = "Welcome"
                };

                var json = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                string baseUrl = _configuration["ApiSettings:BaseUrl"] ?? throw new InvalidOperationException("BaseUrl not configured");
                string registerLoginEndpoint = _configuration["ApiSettings:RegisterLoginEndpoint"] ?? throw new InvalidOperationException("RegisterLoginEndpoint not configured");
                string apiUrl = $"{baseUrl}{registerLoginEndpoint}";

                var response = await httpClient.PostAsync(apiUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    // Add User to UserValidUpto Table

                    UserValidUptoModel user = new UserValidUptoModel
                    {
                        userName = lblFullName,
                        userEmail = lblBussinesseMail,
                        userGstin = lblGSTIN,
                    };

                    await _userBusiness.saveUserValidUpto(user);

                    await _userBusiness.markAsAdmin(lblBussinesseMail);

                    #region Send Email to that User 

                    string toEmail = _configuration["Mail:ToMail"];
                    if (_configuration["Mail:SendToClient"] == "Yes")
                    {
                        toEmail = lblBussinesseMail;
                    }
                    string fromMail = _configuration["Mail:FromMail"];

                    string subjectTemplate = _configuration["SignUpNotification:SubjectTemplate"];
                    string bodyTemplate = _configuration["SignUpNotification:BodyTemplate"];


                    string subject = subjectTemplate;

                    string body = bodyTemplate
                            .Replace("{UserName}", lblFullName)
                            .Replace("{Email}", lblBussinesseMail)
                            .Replace("{Phone}", lblPhoneNo)
                            .Replace("{Gstin}", lblGSTIN)
                            .Replace("{LoginLink}", _configuration["loginPage"]);

                    string[] ccList = _configuration["Mail:CCMail"].Split(',');

                    await SendEmailAsync(
                        toEmail,
                        subject,
                        body,
                        ccList
                    );

                    #endregion

                    var responseContent = await response.Content.ReadAsStringAsync();
                    ViewBag.SuccessMessage = "User Registered Successfully"; // Set success message
                    ViewBag.FormValues = formValues; // Ensure form values are available for initial render
                    return View("Signup");
                }
                else
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    var errorData = JsonConvert.DeserializeObject<dynamic>(errorContent);
                    string errorMessage = errorData?.errorMessage?.ToString() ?? "An error occurred during registration.";
                    ViewBag.ErrMessages = errorMessage;
                    return View("Signup");
                }
            }

            // If inputs are null, show the signup page with a message
            ViewBag.Messages = "RegisterPage";
            return View("Signup");
        }

        public async Task SendEmailAsync(string toEmail, string subject, string body, string[] ccEmails)
        {

            try
            {
                var fromEmail = _configuration["Mail:FromMail"];
                var password = _configuration["Mail:FromMailPassword"];
                var smtpHost = _configuration["Mail:SmtpHost"];
                var smtpPort = int.Parse(_configuration["Mail:SmtpPort"]);
                var enableSsl = bool.Parse(_configuration["Mail:EnableSSL"]);


                MailMessage mail = new MailMessage
                {
                    From = new MailAddress(fromEmail),
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = true
                };

                mail.To.Add(toEmail);

                if (ccEmails != null)
                {
                    foreach (var cc in ccEmails)
                    {
                        if (!string.IsNullOrWhiteSpace(cc))
                            mail.CC.Add(cc.Trim());
                        //continue;
                    }
                }

                using (SmtpClient smtp = new SmtpClient(smtpHost, smtpPort))
                {
                    smtp.Credentials = new NetworkCredential(fromEmail, password);
                    smtp.EnableSsl = enableSsl;
                    await smtp.SendMailAsync(mail);
                }
            }
            catch (Exception ex)
            {
                // Log error or handle appropriately
                Console.WriteLine($"Email sending failed: {ex.Message}");
                throw;
            }
        }

        #endregion

        #region CheckSession
        [HttpGet]
        public IActionResult CheckSession()
        {
            bool isAuthenticated = HttpContext.Session.TryGetValue("UserId", out _);
            return Json(new { isAuthenticated });
        }
        #endregion

        #region TjCaptions
        //public async Task<IActionResult> TjCaptions(string screenName)
        //{
        //    var username = HttpContext.Session.GetString("UserName") ?? "na";
        //    //_logger.LogInformation("TjCaptions called with screenName: {ScreenName}", screenName);

        //    if (string.IsNullOrEmpty(screenName))
        //    {
        //        ViewBag.tjMessage = "Screen Name is required.";
        //        return View();
        //    }

        //    try
        //    {
        //        var httpClient = _httpClientFactory.CreateClient(); // Create HttpClient
        //        var captions = await CommonHelper.GetCaptionsAsync(
        //            screenName,
        //            username,
        //            _logger,
        //            _configuration,
        //            httpClient); // Pass HttpClient

        //        if (captions != null)
        //        {
        //            ViewBag.ResponseDict = captions;
        //            if (!captions.ContainsKey("txtWelcomeBack"))
        //            {
        //                captions["txtWelcomeBack"] = "Welcome Back!";
        //            }
        //        }
        //        else
        //        {
        //            ViewBag.tjMessage = "Invalid ScreenName or error fetching captions.";
        //            //_logger.LogWarning("Captions returned null");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        //_logger.LogError(ex, "Error fetching captions for {ScreenName}", screenName);
        //        ViewBag.tjMessage = "Error fetching captions";
        //    }

        //    return View();
        //}

        public async Task<IActionResult> TjCaptions(string screenName)
        {
            var username = HttpContext.Session.GetString("UserName") ?? "na";

            if (string.IsNullOrEmpty(screenName))
            {
                ViewBag.tjMessage = "Screen Name is required.";
                return View();
            }

            try
            {
                var httpClient = _httpClientFactory.CreateClient();
                var result = await CommonHelper.GetCaptionsAsync(screenName, username, _logger, _configuration, httpClient);

                if (result.Captions != null)
                {
                    ViewBag.ResponseDict = result.Captions;
                    if (!result.Captions.ContainsKey("txtWelcomeBack"))
                    {
                        result.Captions["txtWelcomeBack"] = "Welcome Back!";
                    }
                }
                else
                {
                    ViewBag.tjMessage = $"{result.ErrorMessage}";
                }
            }
            catch (Exception ex)
            {
                ViewBag.tjMessage = $"Exception occurred: {ex.Message}";
            }

            return View();
        }

        #endregion

        #region User Model
        public class User
        {
            public string userName { get; set; }

            public string firstName { get; set; }
            public string emailId { get; set; }
            public string userCode { get; set; }
            public string gstin { get; set; }
            public string password { get; set; }
            public string passwordChanged { get; set; }
        }

        #endregion

    }
}