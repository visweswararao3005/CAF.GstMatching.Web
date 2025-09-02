using Microsoft.AspNetCore.Mvc;
using CAF.GstMatching.Web.Helpers;
using CAF.GstMatching.Web.Common;
using System.Text;
using System.Data;
using DataTable = System.Data.DataTable;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Globalization;
using CAF.GstMatching.Web.Models;
using System.Net.Mail;
using System.Net;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using CellType = NPOI.SS.UserModel.CellType;
using System.Text.Json;
using System.IdentityModel.Tokens.Jwt;
using CAF.GstMatching.Web.Hubs;
using Microsoft.AspNetCore.SignalR;
using CAF.GstMatching.Business.Interface;
using CAF.GstMatching.Models;
using CAF.GstMatching.Models.CompareGst;
using System.Security.Cryptography;
using System.Linq;
using System.Net.Http;
using Org.BouncyCastle.Asn1.Ocsp;
using System.IO;
using System.IO.Compression;
using SharpCompress.Archives;
using SharpCompress.Archives.Tar;
using SharpCompress.Readers;
using SharpCompress.Common;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Diagnostics;
using CAF.GstMatching.Models.UserModel;
using Org.BouncyCastle.Bcpg;
using static CAF.GstMatching.Web.Controllers.HomeController;
using CAF.GstMatching.Models.PurchaseTicketModel;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using Org.BouncyCastle.Utilities.Net;
using System.Net.Sockets;
using Razorpay.Api;

namespace CAF.GstMatching.Web.Controllers
{
    public class AdminController : Controller
	{
		private readonly ILogger<AdminController> _logger;
        private readonly HttpClient _httpClient;
        private readonly IHttpClientFactory _httpClientFactory;
        private readonly IConfiguration _configuration;
        private readonly IHubContext<ChatHub> _hubContext;


        private readonly IUserBusiness _userBusiness;

        private readonly IPurchaseDataBusiness _purchaseDataBusiness;
        private readonly IPurchaseTicketBusiness _purchaseTicketBusiness;	
		private readonly IGSTR2DataBusiness _gSTR2DataBusiness;
		private readonly ICompareGstBusiness _compareGstBusiness;
		private readonly IModifiedDataBusiness _ModifiedDataBusiness;

		private readonly ISLDataBusiness _sLDataBusiness;
		private readonly ISLTicketsBusiness _sLTicketsBusiness;
		private readonly ISLEInvoiceBusiness _sLEInvoiceBusiness;
		private readonly ISLEWayBillBusiness _sLEWayBillBusiness;
		private readonly ISLComparedDataBusiness _sLComparedDataBusiness;

        private readonly INoticeDataBusiness _noticeDataBusiness;

		public AdminController(ILogger<AdminController> logger,
                               IHttpClientFactory httpClientFactory,
                               IConfiguration configuration,
                               IHubContext<ChatHub> hubContext,

                               IUserBusiness userBusiness,

                               IPurchaseDataBusiness purchaseDataBusiness,
							   IPurchaseTicketBusiness purchaseTicketBusiness,
							   IGSTR2DataBusiness gSTR2DataBusiness,
                               ICompareGstBusiness compareGstBusiness,
                               IModifiedDataBusiness ModifiedDataBusiness,
							   
							   ISLDataBusiness sLDataBusiness,
							   ISLTicketsBusiness sLTicketsBusiness,
							   ISLEInvoiceBusiness sLEInvoiceBusiness,
							   ISLEWayBillBusiness sLEWayBillBusiness,
							   ISLComparedDataBusiness sLComparedDataBusiness,

                               INoticeDataBusiness noticeDataBusiness                 
            )

		{
			_logger = logger;
            _httpClientFactory = httpClientFactory;
            _configuration = configuration;
            _hubContext = hubContext;

            _userBusiness = userBusiness;

            _purchaseDataBusiness = purchaseDataBusiness;
            _purchaseTicketBusiness = purchaseTicketBusiness;
			_gSTR2DataBusiness = gSTR2DataBusiness;
			_compareGstBusiness = compareGstBusiness;
			_ModifiedDataBusiness = ModifiedDataBusiness;

			_sLDataBusiness = sLDataBusiness;
			_sLTicketsBusiness = sLTicketsBusiness;
			_sLEInvoiceBusiness = sLEInvoiceBusiness;
			_sLEWayBillBusiness = sLEWayBillBusiness;
			_sLComparedDataBusiness = sLComparedDataBusiness;

            _noticeDataBusiness = noticeDataBusiness;

        }

        #region MapGSTIN
        public async Task<IActionResult> MapGstin()
        {
            var allClients = await _userBusiness.GetAllClients();
			//Console.WriteLine($"All Clients Count: {allClients.Count}");
			//foreach (var client in allClients)
			//{
			//	Console.WriteLine($"Client Data: {client.ClientGSTIN} - {client.ClientName}");
			//}

			string email = MySession.Current.Email; // Get the email from session if not provided         
            var mappedClients = await _userBusiness.GetAdminClients(email);
			//Console.WriteLine($"Mapped Clients Count: {mappedClients.Count}");
			//foreach (var client in mappedClients)
			//{
			//	Console.WriteLine($"Client Data: {client.ClientGSTIN} - {client.ClientName}");
			//}

			var unmappedClients = allClients.Except(mappedClients).ToList();
			//Console.WriteLine($"Unmapped Clients Count: {unmappedClients.Count}");
			//foreach (var client in unmappedClients)
			//{
			//	Console.WriteLine($"Client Data: {client.ClientGSTIN} - {client.ClientName}");
			//}

			ViewBag.mappedClients = mappedClients;
            ViewBag.unmappedClients = unmappedClients;
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act7"; // For active button styling
            //ViewBag.isServer = _configuration["IsServer"] == "true" ? true : false;
            //ViewBag.serverURl = _configuration["ServerUrl"];

            return View("~/Views/Admin/MapGstin/MapGstin.cshtml");

        }

        public async Task<IActionResult> AddClientsOrRemoveClients([FromBody] ClientActionPayload payload)
		{
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act7"; // For active button styling
            
			string email = MySession.Current.Email;
            string flag = payload.Flag;
            var clients = payload.Clients;
			//Console.WriteLine("Email: " + email); // ✅
			//Console.WriteLine("flag: " + flag); // ✅
			Console.WriteLine("Clients Count: " + clients.Count); // ✅
			foreach (var client in clients)
			{
				Console.WriteLine($"Client Data: {client.ClientGSTIN} - {client.ClientName}");
			}
            //Console.ReadKey();

            try
            {
                await _userBusiness.MapClientToAdmin(email, flag, clients);
            }
            catch (Exception ex)
            {
                // You can log the exception or handle it as needed
                // Example: log the error
                _logger.LogError(ex, "Error occurred while mapping client to admin for email: {Email}", email);
                return StatusCode(500, "An error occurred on the server.");
            }

            return RedirectToAction("MapGstin", "Admin");
        }

        // Model to deserialize the JSON payload
        public class ClientActionPayload
        {
            public string Flag { get; set; }
            public List<AdminClientModel> Clients { get; set; }
        }

        #endregion

        #region Manage User Access
        public async Task<IActionResult> ManageUserAccess()
        {
            var userEmail = HttpContext.Session.GetString("Email") ?? "";

            // Get the list of admin emails from configuration
            var adminEmails = (await _userBusiness.GetMainAdminUsers())
									.Select(e => e.Email.ToLower())
									.ToArray();

			// Restrict access if user not in admin list
			if (!adminEmails.Contains(userEmail, StringComparer.OrdinalIgnoreCase))
            {
                return Unauthorized(); // or RedirectToAction("AccessDenied") or 403 view
            }

            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act8";
            var allUsers = await _userBusiness.GetAllUsers();
            // Filter out users whose Email is in adminEmails
            var filteredUsers = allUsers
                .Where(u => !adminEmails.Contains(u.Email?.Trim().ToLower()))
                .ToList();
            ViewBag.allUsers = filteredUsers;

            return View("~/Views/Admin/ManageUserAccess/ManageUserAccess.cshtml");
        }

		public async Task<string> GetUserAccessValidDate(string email)
		{
			var user = await _userBusiness.getUserValidUpto(email);
			return user?.accessToDate?.ToString("dd-MM-yyyy") ?? "";
		}

		[HttpPost]
        public async Task<IActionResult> UpdateUserAccess(string SelectedUser, DateTime NewValidDate, string selectedGstin)
        {
			UserValidUptoModel saveuseraccess = new UserValidUptoModel
			{
				userEmail = SelectedUser,
				accessToDate = NewValidDate,
				userGstin = selectedGstin
            };
            await _userBusiness.saveUserValidUpto(saveuseraccess);
            ViewBag.Message = "New Access date updated successfully.";
			// Get the list of admin emails from configuration
			var adminEmails = (await _userBusiness.GetMainAdminUsers())
									 .Select(e => e.Email.ToLower())
									 .ToArray();
			var allUsers = await _userBusiness.GetAllUsers();
            // Filter out users whose Email is in adminEmails
            var filteredUsers = allUsers
                .Where(u => !adminEmails.Contains(u.Email?.Trim().ToLower()))
                .ToList();
            ViewBag.allUsers = filteredUsers;
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act8";
            return View("~/Views/Admin/ManageUserAccess/ManageUserAccess.cshtml");
        }

        #endregion

        #region TjCaptions
        private async Task TjCaptions(string screenName)
		{
			var username = "na"; // Replace with actual username if needed
			var httpClient = _httpClientFactory.CreateClient();
			var captions = await CommonHelper.GetCaptionsAsync(screenName, username, _logger, _configuration, httpClient);
			ViewBag.ResponseDict = captions;
		}
        #endregion

        #region Purchase Register Upload Invoice File
        // ✅ Compare GST Page
        public async Task<IActionResult> EditUploadGST(string requestNo)
        {
            var clientGSTIN = MySession.Current.gstin;
            var ticketDetails = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(requestNo, clientGSTIN);
            ViewBag.RequestNo = requestNo;
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            ViewBag.FinancialYear = ticketDetails.FinancialYear;
            ViewBag.PeriodType = ticketDetails.PeriodType;
            ViewBag.TxnPeriod = ticketDetails.TxnPeriod;
            ViewBag.FileName = ticketDetails.FileName;
            return View("~/Views/Admin/PurchaseRegister/UploadGST.cshtml");
        }
        public IActionResult Upload(string ticketId, string Edit)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            ViewBag.Message = ticketId;
            ViewBag.Edit = Edit;
            return View("~/Views/Admin/PurchaseRegister/UploadGST.cshtml");
        }
        [HttpPost]
        public async Task<IActionResult> UploadGST(IFormFile gstFile, string financialYear, string periodtype, string period, string requestNo)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            ViewBag.requestNo = requestNo;
            ViewBag.Edit = string.IsNullOrEmpty(requestNo) ? "No" : "Yes";
            string[] validMonths = null;
            // Determine valid months based on periodType
            if (periodtype.Equals("Monthly", StringComparison.OrdinalIgnoreCase))
            {
                validMonths = new string[] { period }; // period in Nov-24

            }
            else // Quarterly  Q1-2025, Q2-2025, Q3-2025, Q4-2026  for FY 2025-26
            {
                var split = period.Split('-');

                string quarter = split[0]; // Q1,Q2,Q3,Q4
                string year = split[1];    // 2025,2025,2025,2026
                string shortYear = year.Substring(2); // "25","25", "25", "26"  


                switch (quarter.ToUpper())
                {
                    case "Q1":
                        validMonths = new string[] { $"Apr-{shortYear}", $"May-{shortYear}", $"Jun-{shortYear}" };
                        break;
                    case "Q2":
                        validMonths = new string[] { $"Jul-{shortYear}", $"Aug-{shortYear}", $"Sep-{shortYear}" };
                        break;
                    case "Q3":
                        validMonths = new string[] { $"Oct-{shortYear}", $"Nov-{shortYear}", $"Dec-{shortYear}" };
                        break;
                    case "Q4":
                        validMonths = new string[] { $"Jan-{shortYear}", $"Feb-{shortYear}", $"Mar-{shortYear}" };
                        break;
                    default:
                        throw new Exception($"Invalid quarter: '{quarter}'. Expected Q1 to Q4.");
                }
            }
            if (gstFile == null || gstFile.Length == 0)
            {

                ViewBag.ErrorMessage = "Please select a valid file.";
                return View("~/Views/Admin/PurchaseRegister/UploadGST.cshtml");
            }
            DataTable dataTable;
            string extension = Path.GetExtension(gstFile.FileName);
            string fileName = Path.GetFileName(gstFile.FileName);
            var _userGstin = MySession.Current.gstin; // Get the GSTIN from the session

            if (extension == ".csv")
            {
                using (var stream = new MemoryStream())
                {
                    await gstFile.CopyToAsync(stream);
                    stream.Position = 0;
                    try
                    {
                        dataTable = ReadCsvFile(stream, validMonths, _userGstin);
                        if (!ValidateColumnNamesInvoiceUpload(dataTable))
                        {
                            ViewBag.ErrorMessage = "Please check columns names";
                            return View("~/Views/Admin/PurchaseRegister/UploadGST.cshtml");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        return View("~/Views/Admin/PurchaseRegister/UploadGST.cshtml");
                    }
                }

            }
            else if (extension == ".xlsx")
            {
                string sheetName = _configuration["PR_Invoice_Xlsx_SheetName"];
                using (var stream = new MemoryStream())
                {
                    await gstFile.CopyToAsync(stream);
                    stream.Position = 0;
                    try
                    {
                        dataTable = ReadExcelFile(stream, validMonths, sheetName, _userGstin); //Change sheet name from hard-code to get from user  
                        if (!ValidateColumnNamesInvoiceUpload(dataTable))
                        {
                            ViewBag.ErrorMessage = "Please check columns names";
                            return View("~/Views/Admin/PurchaseRegister/UploadGST.cshtml");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        return View("~/Views/Admin/PurchaseRegister/UploadGST.cshtml");
                    }
                }
            }
            else if (extension == ".xls")
            {
                string sheetName = _configuration["PR_Invoice_Xlsx_SheetName"];
                using (var stream = new MemoryStream())
                {
                    await gstFile.CopyToAsync(stream);
                    stream.Position = 0;
                    try
                    {
                        dataTable = ReadXLSFile(stream, validMonths, sheetName, _userGstin); //Change sheet name from hard-code to get from user  
                        if (!ValidateColumnNamesInvoiceUpload(dataTable))
                        {
                            ViewBag.ErrorMessage = "Please check columns names";
                            return View("~/Views/Admin/PurchaseRegister/UploadGST.cshtml");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        return View("~/Views/Admin/PurchaseRegister/UploadGST.cshtml");
                    }
                }
            }
            else
            {
                ViewBag.ErrorMessage = "Invalid file format. Please upload CSV/Xlsx/Xls file";
                return View("~/Views/Admin/PurchaseRegister/UploadGST.cshtml");
            }

            //_logger.LogInformation("rows in datatable: {0}", dataTable.Rows.Count); 
            // Generate a ticket

            string name = MySession.Current.UserName;
            string usergstin = MySession.Current.gstin;
            ViewBag.ticketId = string.IsNullOrEmpty(requestNo) ? GenerateTicket() : requestNo;
            string ticketId = ViewBag.ticketId;
            try
            {
                await _purchaseDataBusiness.SavePurchaseDataAsync(dataTable, ticketId, usergstin);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"{ex.Message} ";
                return View("~/Views/Admin/PurchaseRegister/UploadGST.cshtml");
            }

            string Edit = ViewBag.Edit;
            DateTime? createdDate = DateTime.Now;
            if (ViewBag.Edit == "Yes")
            {
                var clientGSTIN = MySession.Current.gstin;
                var ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(requestNo, clientGSTIN);
                createdDate = ticket.RequestCreatedDate;
            }

            TicketsStatusModel data = new TicketsStatusModel();
            {
                data.EXERTUSERNAME = name;
                data.ClientGSTIN = usergstin;
                data.RequestNo = ticketId;
                data.RequestUpdatedDate = DateTime.Now;
                data.RequestCreatedDate = createdDate;
                data.FinancialYear = financialYear;
                data.PeriodType = periodtype;
                data.TxnPeriod = period;
                data.FileName = fileName;
                data.Email = MySession.Current.Email;
            }

            await _purchaseTicketBusiness.SavePurchaseTicketAsync(data);

            return RedirectToAction("Upload", "Admin", new { ticketId, Edit });
        }
        private string getNumber(string str)
        {
            switch (str)
            {
                case "January":
                case "Jan":
                    return "01";
                case "February":
                case "Feb":
                    return "02";
                case "March":
                case "Mar":
                    return "03";
                case "April":
                case "Apr":
                    return "04";
                case "May":

                    return "05";
                case "June":
                case "Jun":
                    return "06";
                case "July":
                case "Jul":
                    return "07";
                case "August":
                case "Aug":
                    return "08";
                case "September":
                case "Sep":
                    return "09";
                case "October":
                case "Oct":
                    return "10";
                case "November":
                case "Nov":
                    return "11";
                case "December":
                case "Dec":
                    return "12";
                default:
                    throw new ArgumentException("Invalid month name");
            }
        }
        // Function to Generate a Request ID
        private string GenerateTicket()
        {
            return "REQ_PR_" + DateTime.Now.ToString("yyyyMMddHHmmss");
        }
        private DataTable ReadCsvFile(Stream stream, string[] period, string gstin)
        {
            var _invoice = _configuration["Invoice"];
            var _invoiceColumns = _invoice.Split(',').Select(x => x.Trim()).ToList();

            DataTable dt = new DataTable();
            int lineNumber = 0;

            int gstinIndex = -1;
            int periodColumnIndex = -1;
            int supplierGstinIndex = -1;
            int supplierNameIndex = -1;
            int invoiceNoColumnIndex = -1;
            int invoiceDateColumnIndex = -1;

            List<int> columnMismatchRows = new List<int>();
            List<int> gstinInvalidRows = new List<int>();
            List<int> periodMismatchRows = new List<int>();
            List<int> supplierGstinInvalidRows = new List<int>();
            List<int> supplierNameInvalidRows = new List<int>();
            List<int> invoiceNoInvalidRows = new List<int>();
            List<int> invoiceDateMismatchRows = new List<int>();


            string[] PeriodSub = period.Select(p => p.Substring(0, 3)).ToArray();
            string[] expectednumPeriod = period.Select(p => getNumber(p.Substring(0, 3))).ToArray();
            string[] periodYear = period.Select(p => p.Substring(p.Length - 2)).ToArray();
            string UserGstin = gstin; // Remove all internal spaces and trim
            using (var reader = new StreamReader(stream))
            using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(reader))
            {
                parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = true;

                // Read header
                string[] headers = parser.ReadFields();
                lineNumber++;

                foreach (var header in headers)
                    dt.Columns.Add(header.Trim(), typeof(string));
                //User GSTIN period Supplier GSTIN Supplier Name invoiceno invoice_date Taxable Value  cgst sgst Integrated Tax  Cess

                // ✅ Gstin validation
                gstinIndex = Array.FindIndex(headers, h => h.Equals(_invoiceColumns[0], StringComparison.OrdinalIgnoreCase));
                if (gstinIndex == -1)
                    throw new Exception($"The Excel file does not contain a '{_invoiceColumns[0]}' column.");
                // Period validation
                periodColumnIndex = Array.FindIndex(headers, h => h.Trim().Equals(_invoiceColumns[1], StringComparison.OrdinalIgnoreCase));
                if (periodColumnIndex == -1)
                    throw new Exception($"The CSV file does not contain a '{_invoiceColumns[1]}' column.");
                // Supplier GSTIN validation
                supplierGstinIndex = Array.FindIndex(headers, h => h.Trim().Equals(_invoiceColumns[2], StringComparison.OrdinalIgnoreCase));
                if (supplierGstinIndex == -1)
                    throw new Exception($"The CSV file does not contain a '{_invoiceColumns[2]}' column.");
                // Supplier name validation
                supplierNameIndex = Array.FindIndex(headers, h => h.Trim().Equals(_invoiceColumns[3], StringComparison.OrdinalIgnoreCase));
                if (supplierNameIndex == -1)
                    throw new Exception($"The CSV file does not contain a '{_invoiceColumns[3]}' column.");
                // Invoice number validation
                invoiceNoColumnIndex = Array.FindIndex(headers, h => h.Trim().Equals(_invoiceColumns[4], StringComparison.OrdinalIgnoreCase));
                if (invoiceNoColumnIndex == -1)
                    throw new Exception($"The CSV file does not contain a '{_invoiceColumns[4]}' column.");
                // Invoice date validation
                invoiceDateColumnIndex = Array.FindIndex(headers, h => h.Trim().Equals(_invoiceColumns[5], StringComparison.OrdinalIgnoreCase));
                if (invoiceDateColumnIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{_invoiceColumns[5]}' column.");

                // Read data rows
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    lineNumber++;
                    // Column count check
                    if (fields.Length != dt.Columns.Count)
                    {
                        columnMismatchRows.Add(lineNumber);
                        continue;
                    }
                    // GSTIN Validation
                    string gstinValue = fields[gstinIndex];
                    if (string.IsNullOrWhiteSpace(gstinValue) || gstinValue.Length != 15 || gstinValue != UserGstin)
                    {
                        gstinInvalidRows.Add(lineNumber);
                        continue;
                    }
                    //Period validation
                    string rowPeriod = fields[periodColumnIndex]?.Trim().ToLower();
                    //Console.WriteLine($"rowPeriod: {rowPeriod}");
                    //if (string.IsNullOrEmpty(rowPeriod))
                    //{
                    //    periodMismatchRows.Add(lineNumber);
                    //    continue;
                    //}
                    // Supplier GSTIN Validation
                    string supplierGstinValue = fields[supplierGstinIndex];
                    if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
                    {
                        supplierGstinInvalidRows.Add(lineNumber);
                        continue;
                    }
                    // Supplier Name Validation
                    string supplierNameValue = fields[supplierNameIndex];
                    if (supplierNameValue.Length > 100)
                    {
                        supplierNameInvalidRows.Add(lineNumber);
                        continue;
                    }
                    //invoice number Validation
                    string invoiceNoValue = fields[invoiceNoColumnIndex];
                    if (string.IsNullOrWhiteSpace(invoiceNoValue) || invoiceNoValue.Length > 25)
                    {
                        invoiceNoInvalidRows.Add(lineNumber);
                        continue;
                    }
                    //string rowPeriodSub = rowPeriod.Substring(0, 3); // e.g., jan2025 => jan
                    // bool isValidPeriod = PeriodSub.Any(month => month.StartsWith(rowPeriodSub, StringComparison.OrdinalIgnoreCase));
                    string invoiceDateStr = fields[invoiceDateColumnIndex]?.Trim();
                    //bool isValidInvoiceDate = false;
                    string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
                    if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
                    {
                        fields[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
                    }
                    else
                    {
                        //Console.WriteLine(fields[invoiceDateColumnIndex]);
                        invoiceDateMismatchRows.Add(lineNumber);
                        continue;
                    }

                    dt.Rows.Add(fields);

                }
            }

            // Handle mismatch errors
            if (columnMismatchRows.Count > 0 || gstinInvalidRows.Count > 0 || periodMismatchRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || supplierNameInvalidRows.Count > 0 || invoiceNoInvalidRows.Count > 0 || invoiceDateMismatchRows.Count > 0)
            {
                var errorMsg = "";

                if (columnMismatchRows.Count > 0)
                    errorMsg += $"Column count mismatch at line(s): {string.Join(", ", columnMismatchRows)}. Expected column count: {dt.Columns.Count}.";
                if (gstinInvalidRows.Count > 0)
                    errorMsg += $"\n Invalid User GSTIN at line(s): {string.Join(", ", gstinInvalidRows)}.";
                if (periodMismatchRows.Count > 0)
                    errorMsg += $"\n Period is null at line(s): {string.Join(", ", periodMismatchRows)}. Expected month(s): {string.Join(", ", period)}.";
                if (supplierGstinInvalidRows.Count > 0)
                    errorMsg += $"\n Invalid Supplier GSTIN at line(s): {string.Join(", ", supplierGstinInvalidRows)}.";
                if (supplierNameInvalidRows.Count > 0)
                    errorMsg += $"\n The Supplier Name field allows a maximum of 100 characters.Invalid Supplier Name at line(s) : {string.Join(", ", supplierNameInvalidRows)}.";
                if (invoiceNoInvalidRows.Count > 0)
                    errorMsg += $"\n The Invoice Number field allows a maximum of 25 characters.Invalid Invoice Number at line(s) : {string.Join(", ", invoiceNoInvalidRows)}.";
                if (invoiceDateMismatchRows.Count > 0)
                    errorMsg += $"\n Invoice date format mismatch at line(s): {string.Join(", ", invoiceDateMismatchRows)}. Expected Format : dd-MM-yyyy hh:mm:ss tt .";

                throw new Exception("Failed to insert data: \n" + errorMsg);
            }

            return dt;
        }
        private DataTable ReadExcelFile(Stream stream, string[] period, string sheetName, string gstin)
        {
            var _invoice = _configuration["Invoice"];
            var _invoiceColumns = _invoice.Split(',').Select(x => x.Trim()).ToList();

            DataTable dt = new DataTable();
            int lineNumber = 1; // Start at 1 for header
            int gstinIndex = -1;
            int periodColumnIndex = -1;
            int supplierGstinIndex = -1;
            int supplierNameIndex = -1;
            int invoiceNoColumnIndex = -1;
            int invoiceDateColumnIndex = -1;

            List<int> columnMismatchRows = new List<int>();
            List<int> gstinInvalidRows = new List<int>();
            List<int> periodMismatchRows = new List<int>();
            List<int> supplierGstinInvalidRows = new List<int>();
            List<int> supplierNameInvalidRows = new List<int>();
            List<int> invoiceNoInvalidRows = new List<int>();
            List<int> invoiceDateMismatchRows = new List<int>();

            string[] PeriodSub = period.Select(p => p.Substring(0, 3)).ToArray();
            string[] expectednumPeriod = period.Select(p => getNumber(p.Substring(0, 3))).ToArray();
            string[] periodYear = period.Select(p => p.Substring(p.Length - 2)).ToArray();
            using (var workbook = new XLWorkbook(stream))
            {
                //var worksheet = workbook.Worksheets.First();
                var worksheet = workbook.Worksheet(sheetName);
                if (worksheet == null)
                    throw new Exception($"Sheet '{sheetName}' not found.");
                var rows = worksheet.RowsUsed().ToList();

                if (rows.Count == 0)
                    throw new Exception("Excel file is empty.");

                // Read header
                var headerRow = rows[0];
                var headers = headerRow.Cells().Select(c => c.GetString().Trim()).ToArray();

                foreach (var header in headers)
                    dt.Columns.Add(header, typeof(string));
                //User GSTIN period Supplier GSTIN Supplier Name invoiceno invoice_date Taxable Value  cgst sgst Integrated Tax  Cess

                // ✅ Gstin validation
                gstinIndex = Array.FindIndex(headers, h => h.Equals(_invoiceColumns[0], StringComparison.OrdinalIgnoreCase));
                if (gstinIndex == -1)
                    throw new Exception($"The Excel file does not contain a '{_invoiceColumns[0]}' column.");
                // Period validation
                periodColumnIndex = Array.FindIndex(headers, h => h.Trim().Equals(_invoiceColumns[1], StringComparison.OrdinalIgnoreCase));
                if (periodColumnIndex == -1)
                    throw new Exception($"The CSV file does not contain a '{_invoiceColumns[1]}' column.");
                // Supplier GSTIN validation
                supplierGstinIndex = Array.FindIndex(headers, h => h.Trim().Equals(_invoiceColumns[2], StringComparison.OrdinalIgnoreCase));
                if (supplierGstinIndex == -1)
                    throw new Exception($"The CSV file does not contain a '{_invoiceColumns[2]}' column.");
                // Supplier name validation
                supplierNameIndex = Array.FindIndex(headers, h => h.Trim().Equals(_invoiceColumns[3], StringComparison.OrdinalIgnoreCase));
                if (supplierNameIndex == -1)
                    throw new Exception($"The CSV file does not contain a '{_invoiceColumns[3]}' column.");
                // Invoice number validation
                invoiceNoColumnIndex = Array.FindIndex(headers, h => h.Trim().Equals(_invoiceColumns[4], StringComparison.OrdinalIgnoreCase));
                if (invoiceNoColumnIndex == -1)
                    throw new Exception($"The CSV file does not contain a '{_invoiceColumns[4]}' column.");
                // Invoice date validation
                invoiceDateColumnIndex = Array.FindIndex(headers, h => h.Trim().Equals(_invoiceColumns[5], StringComparison.OrdinalIgnoreCase));
                if (invoiceDateColumnIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{_invoiceColumns[5]}' column.");



                // Read data rows
                for (int i = 1; i < rows.Count; i++)
                {
                    int excelLineNumber = i + 1; // Line number as seen in Excel (header is line 1)
                    var row = rows[i];
                    var cells = row.Cells().Select((c, index) =>
                    {
                        if (index == invoiceDateColumnIndex) // Replace with the correct index for invoice_date
                        {
                            if (DateTime.TryParse(c.GetValue<string>(), out DateTime parsedDate))
                            {
                                return parsedDate.ToString("dd-MM-yyyy hh:mm:ss tt"); // or "yyyy - MM - dd" based on your DB requirement
                            }
                        }
                        return c.GetFormattedString()?.Trim() ?? "";
                    }).ToArray();

                    // Column count check
                    if (cells.Length != dt.Columns.Count)
                    {
                        columnMismatchRows.Add(excelLineNumber);
                        continue;
                    }
                    // GSTIN Validation
                    string gstinValue = cells[gstinIndex]; // Remove all internal spaces and trim
                    if (string.IsNullOrWhiteSpace(gstinValue) || gstinValue.Length != 15 || gstinValue != gstin)
                    {
                        gstinInvalidRows.Add(excelLineNumber);
                        continue;
                    }
                    // Period validation
                    string rowPeriod = cells[periodColumnIndex]?.ToLower();
                    //if (string.IsNullOrEmpty(rowPeriod))
                    //{
                    //    periodMismatchRows.Add(excelLineNumber);
                    //    continue;
                    //}
                    // Supplier GSTIN Validation
                    string supplierGstinValue = cells[supplierGstinIndex];
                    if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
                    {
                        supplierGstinInvalidRows.Add(excelLineNumber);
                        continue;
                    }
                    // Supplier Name Validation
                    string supplierNameValue = cells[supplierNameIndex];
                    if (supplierNameValue.Length > 100)
                    {
                        supplierNameInvalidRows.Add(excelLineNumber);
                        continue;
                    }
                    // Invoice number Validation
                    string invoiceNoValue = cells[invoiceNoColumnIndex];
                    if (string.IsNullOrWhiteSpace(invoiceNoValue) || invoiceNoValue.Length > 25)
                    {
                        invoiceNoInvalidRows.Add(excelLineNumber);
                        continue;
                    }
                    // Invoice date Validation
                    string invoiceDateStr = cells[invoiceDateColumnIndex]?.Trim();
                    string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
                    if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
                    {
                        cells[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
                    }
                    else
                    {
                        //Console.WriteLine(fields[invoiceDateColumnIndex]);
                        invoiceDateMismatchRows.Add(lineNumber);
                        continue;
                    }


                    dt.Rows.Add(cells);
                }

            }

            // Handle mismatch errors
            if (columnMismatchRows.Count > 0 || gstinInvalidRows.Count > 0 || periodMismatchRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || supplierNameInvalidRows.Count > 0 || invoiceNoInvalidRows.Count > 0 || invoiceDateMismatchRows.Count > 0)
            {
                var errorMsg = "";

                if (columnMismatchRows.Count > 0)
                    errorMsg += $"Column count mismatch at line(s): {string.Join(", ", columnMismatchRows)}. Expected column count: {dt.Columns.Count}.";
                if (gstinInvalidRows.Count > 0)
                    errorMsg += $"\n Invalid User GSTIN at line(s): {string.Join(", ", gstinInvalidRows)}.";
                if (periodMismatchRows.Count > 0)
                    errorMsg += $"\n Period is null at line(s): {string.Join(", ", periodMismatchRows)}. Expected month(s): {string.Join(", ", period)}.";
                if (supplierGstinInvalidRows.Count > 0)
                    errorMsg += $"\n Invalid Supplier GSTIN at line(s): {string.Join(", ", supplierGstinInvalidRows)}.";
                if (supplierNameInvalidRows.Count > 0)
                    errorMsg += $"\n The Supplier Name field allows a maximum of 100 characters.Invalid Supplier Name at line(s) : {string.Join(", ", supplierNameInvalidRows)}.";
                if (invoiceNoInvalidRows.Count > 0)
                    errorMsg += $"\n The Invoice Number field allows a maximum of 25 characters.Invalid Invoice Number at line(s) : {string.Join(", ", invoiceNoInvalidRows)}.";
                if (invoiceDateMismatchRows.Count > 0)
                    errorMsg += $"\n Invoice date format mismatch at line(s): {string.Join(", ", invoiceDateMismatchRows)}. Expected Format : dd-MM-yyyy hh:mm:ss tt .";

                throw new Exception("Failed to insert data: \n" + errorMsg);
            }

            return dt;
        }
        private DataTable ReadXLSFile(Stream stream, string[] period, string sheetName, string gstin)
        {
            var _invoice = _configuration["Invoice"];
            var _invoiceColumns = _invoice.Split(',').Select(x => x.Trim()).ToList();

            DataTable dt = new DataTable();
            int lineNumber = 1;

            int gstinIndex = -1;
            int periodColumnIndex = -1;
            int supplierGstinIndex = -1;
            int supplierNameIndex = -1;
            int invoiceNoColumnIndex = -1;
            int invoiceDateColumnIndex = -1;

            List<int> columnMismatchRows = new();
            List<int> gstinInvalidRows = new();
            List<int> periodMismatchRows = new();
            List<int> supplierGstinInvalidRows = new();
            List<int> supplierNameInvalidRows = new();
            List<int> invoiceNoInvalidRows = new();
            List<int> invoiceDateMismatchRows = new();

            string[] PeriodSub = period.Select(p => p.Substring(0, 3)).ToArray();
            string[] expectedNumPeriod = period.Select(p => getNumber(p.Substring(0, 3))).ToArray();
            string[] periodYear = period.Select(p => p.Substring(p.Length - 2)).ToArray();
            string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();

            // Load the .xls workbook using NPOI
            HSSFWorkbook workbook = new HSSFWorkbook(stream);
            ISheet sheet = workbook.GetSheet(sheetName) ?? throw new Exception($"Sheet '{sheetName}' not found.");

            if (sheet.PhysicalNumberOfRows == 0)
                throw new Exception("Excel file is empty.");

            // Read header
            IRow headerRow = sheet.GetRow(0);
            int columnCount = headerRow.LastCellNum;
            string[] headers = new string[columnCount];

            for (int i = 0; i < columnCount; i++)
            {
                headers[i] = headerRow.GetCell(i)?.ToString().Trim() ?? $"Column{i}";
                dt.Columns.Add(headers[i], typeof(string));
            }

            // Index mapping and validation
            gstinIndex = FindColumnIndex(headers, _invoiceColumns[0]);
            periodColumnIndex = FindColumnIndex(headers, _invoiceColumns[1]);
            supplierGstinIndex = FindColumnIndex(headers, _invoiceColumns[2]);
            supplierNameIndex = FindColumnIndex(headers, _invoiceColumns[3]);
            invoiceNoColumnIndex = FindColumnIndex(headers, _invoiceColumns[4]);
            invoiceDateColumnIndex = FindColumnIndex(headers, _invoiceColumns[5]);

            // Process rows
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;

                string[] cells = new string[columnCount];
                for (int j = 0; j < columnCount; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell == null)
                    {
                        cells[j] = string.Empty;
                        continue;
                    }

                    if (j == invoiceDateColumnIndex && cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
                    {
                        DateTime DT = (DateTime)cell.DateCellValue;
                        cells[j] = DT.ToString("dd-MM-yyyy hh:mm:ss tt"); // Match your allowed format
                    }
                    else
                    {
                        cells[j] = cell.ToString().Trim();
                    }
                }


                int excelLineNumber = i + 1;

                // Column count check
                if (cells.Length != dt.Columns.Count)
                {
                    columnMismatchRows.Add(excelLineNumber);
                    continue;
                }

                // GSTIN Validation
                string gstinValue = cells[gstinIndex];
                if (string.IsNullOrWhiteSpace(gstinValue) || gstinValue.Length != 15 || gstinValue != gstin)
                {
                    gstinInvalidRows.Add(excelLineNumber);
                    continue;
                }

                // Supplier GSTIN Validation
                string supplierGstin = cells[supplierGstinIndex];
                if (string.IsNullOrWhiteSpace(supplierGstin) || supplierGstin.Length != 15)
                {
                    supplierGstinInvalidRows.Add(excelLineNumber);
                    continue;
                }

                // Supplier Name Validation
                string supplierName = cells[supplierNameIndex];
                if (supplierName.Length > 100)
                {
                    supplierNameInvalidRows.Add(excelLineNumber);
                    continue;
                }

                // Invoice Number Validation
                string invoiceNo = cells[invoiceNoColumnIndex];
                if (string.IsNullOrWhiteSpace(invoiceNo) || invoiceNo.Length > 25)
                {
                    invoiceNoInvalidRows.Add(excelLineNumber);
                    continue;
                }

                // Invoice Date Validation
                string invoiceDateStr = cells[invoiceDateColumnIndex];


                //Console.WriteLine($"Invoice Date: {cells[invoiceDateColumnIndex]}");
                if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
                {
                    cells[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
                }
                else
                {

                    invoiceDateMismatchRows.Add(excelLineNumber);
                    continue;
                }
                //Console.WriteLine($"M - Invoice Date: {cells[invoiceDateColumnIndex]}");

                // Add row to DataTable
                dt.Rows.Add(cells);
            }

            // Handle mismatch errors
            if (columnMismatchRows.Count > 0 || gstinInvalidRows.Count > 0 || periodMismatchRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || supplierNameInvalidRows.Count > 0 || invoiceNoInvalidRows.Count > 0 || invoiceDateMismatchRows.Count > 0)
            {
                var errorMsg = "";

                if (columnMismatchRows.Count > 0)
                    errorMsg += $"Column count mismatch at line(s): {string.Join(", ", columnMismatchRows)}. Expected column count: {dt.Columns.Count}.";
                if (gstinInvalidRows.Count > 0)
                    errorMsg += $"\n Invalid User GSTIN at line(s): {string.Join(", ", gstinInvalidRows)}.";
                if (periodMismatchRows.Count > 0)
                    errorMsg += $"\n Period is null at line(s): {string.Join(", ", periodMismatchRows)}. Expected month(s): {string.Join(", ", period)}.";
                if (supplierGstinInvalidRows.Count > 0)
                    errorMsg += $"\n Invalid Supplier GSTIN at line(s): {string.Join(", ", supplierGstinInvalidRows)}.";
                if (supplierNameInvalidRows.Count > 0)
                    errorMsg += $"\n The Supplier Name field allows a maximum of 100 characters.Invalid Supplier Name at line(s) : {string.Join(", ", supplierNameInvalidRows)}.";
                if (invoiceNoInvalidRows.Count > 0)
                    errorMsg += $"\n The Invoice Number field allows a maximum of 25 characters.Invalid Invoice Number at line(s) : {string.Join(", ", invoiceNoInvalidRows)}.";
                if (invoiceDateMismatchRows.Count > 0)
                    errorMsg += $"\n Invoice date format mismatch at line(s): {string.Join(", ", invoiceDateMismatchRows)}. Expected Formate : dd-MM-yyyy hh:mm:ss tt .";

                throw new Exception("Failed to insert data: \n" + errorMsg);
            }


            return dt;
        }
        private int FindColumnIndex(string[] headers, string expectedName)
        {
            int index = Array.FindIndex(headers, h => h.Equals(expectedName, StringComparison.OrdinalIgnoreCase));
            if (index == -1)
                throw new Exception($"The XLS file does not contain a '{expectedName}' column.");
            return index;
        }
        private bool ValidateColumnNamesInvoiceUpload(DataTable dataTable)
        {
            var _invoice = _configuration["Invoice"];
            var _invoiceColumns = _invoice.Split(',').Select(x => x.Trim()).ToList();
            var uploadedColumns = dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName.ToLower()).ToList();
            return _invoiceColumns.All(col => uploadedColumns.Contains(col.ToLower()));
        }

        #endregion

        #region PurchaseRegisterCurrentRequestsCSV
        public IActionResult OpenRequestsCSV()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act2"; // For active button styling
			return RedirectToAction("OpenTaskCSV", "Admin");
		}

		public async Task<IActionResult> OpenTaskCSV(DateTime? fromdate, DateTime? todate)
		{

			//Console.WriteLine($"Email : {email}"); // Log the email for debugging
			// Default to 2 days ago and today if no values are provided
			DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
			DateTime toDateTime = todate ?? DateTime.Now; // Default to today

			//await TjCaptions("OpenTasks"); // Load captions for OpenTask page
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act2"; // For active button styling
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                return View("~/Views/Admin/OpenTasks/OpenTaskCSV.cshtml");
            }
            //var Tickets = await _purchaseTicketBusiness.GetAllOpenTicketsStatusAsync(fromDateTime, toDateTime);
            //var email = MySession.Current.Email; // Get the email from session
			//var clients = await _userBusiness.GetAdminClients(email);
			//string[] gstinArray = clients.Select(c => c.ClientGSTIN).ToArray();
			//var Tickets = await _purchaseTicketBusiness.GetClientsOpenTicketsStatusAsync(gstinArray, fromDateTime, toDateTime);
			//Console.WriteLine("Client GSTINs: " + string.Join(", ", gstinArray));

			// HttpContext.Session.SetString("TicketsList", JsonConvert.SerializeObject(Tickets));
			//ViewBag.Tickets = Tickets;


            var Tickets = await _purchaseTicketBusiness.GetOpenTicketsStatusAsync(MySession.Current.gstin, fromDateTime, toDateTime); //(ViewBag.UserGSTIN); // Fetch pending tickets based on user_gstin

            ViewBag.Tickets = Tickets;


            return View("~/Views/Admin/OpenTasks/OpenTaskCSV.cshtml"); // Admin dashboard

			//based on his clients Gstin Fetch tickets and show in opentickets page
			//step -1
			// based on his id get the client gstin list from Admin_client_link table
			//let say this Admin with back@end.com mail id have 3 clients their gstin is {1,2,3} 
			//string[] clients = { "1", "2", "3" }; // create and call function here
			//step -2
			//Based on this clients Gstin fetch Pending the tickets from Purchase_Ticket table
			//var Tickets = await _purchaseTicketBusiness.GetClientTicketsStatusAsync(clients);
		}

		public async Task<IActionResult> CompareGSTCSV(string ticketnumber, string ClientGSTIN, string fromDate, string toDate)
		{
			// _logger.LogInformation($"Ticket Number: {ticketnumber}");
			//based on this ticket number fetch data from Purchase Ticket table and store
			var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);

			ViewBag.Ticket = Ticket;
			ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            ViewBag.fromDate = fromDate;
			ViewBag.toDate = toDate;

			return View("~/Views/Admin/CompareGstFiles/CompareGSTCSV.cshtml");
		}

		public async Task<IActionResult> CompareCSV(IFormFile Portal, string ticketnumber, string ClientGSTIN)
		{
			//    string fileName = Path.GetFileName(Portal.FileName);
			//    ViewBag.FileName = fileName;
			ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
			ViewBag.Ticket = Ticket;

			if (Portal == null || Portal.Length == 0)
			{
				ViewBag.ErrorMessage = "Please select a valid file.";
				return View("/Views/Admin/CompareGstFiles/CompareGSTCSV.cshtml");
			}
			DataTable dataTable;
			string extension = Path.GetExtension(Portal.FileName);
			string AdminFileName = Path.GetFileName(Portal.FileName);
			//Console.WriteLine($"AdminFileName: {AdminFileName}"); // Log the file name for debugging

			if (extension == ".csv")
			{
				using (var stream = new MemoryStream())
				{
					await Portal.CopyToAsync(stream);
					stream.Position = 0;
					try
					{
						dataTable = ReadCsvFile(stream, ClientGSTIN);
						if (!ValidateColumnNames(dataTable))
						{
							ViewBag.ErrorMessage = "Please check columns names";
							return View("/Views/Admin/CompareGstFiles/CompareGSTCSV.cshtml");
						}
					}
					catch (Exception ex)
					{
						ViewBag.ErrorMessage = ex.Message;
						return View("/Views/Admin/CompareGstFiles/CompareGSTCSV.cshtml");
					}
				}
			}
			else if (extension == ".xlsx")
			{
				string sheetName = _configuration["PR_Portal_Xlsx_SheetName"];
				using (var stream = new MemoryStream())
				{
					await Portal.CopyToAsync(stream);
					stream.Position = 0;
					try
					{
						dataTable = ReadExcelFile(stream, sheetName, ClientGSTIN); //Change sheet name from hard-code to get from user  
						if (!ValidateColumnNames(dataTable))
						{
							ViewBag.ErrorMessage = "Please check columns names";
							return View("/Views/Admin/CompareGstFiles/CompareGSTCSV.cshtml");
						}
					}
					catch (Exception ex)
					{
						ViewBag.ErrorMessage = ex.Message;
						return View("/Views/Admin/CompareGstFiles/CompareGSTCSV.cshtml");
					}
				}
			}
			else if (extension == ".xls")
			{
				string sheetName = _configuration["PR_Portal_Xlsx_SheetName"];
				using (var stream = new MemoryStream())
				{
					await Portal.CopyToAsync(stream);
					stream.Position = 0;
					try
					{
						dataTable = ReadXLSFile(stream, sheetName, ClientGSTIN); //Change sheet name from hard-code to get from user  
						if (!ValidateColumnNames(dataTable))
						{
							ViewBag.ErrorMessage = "Please check columns names";
							return View("/Views/Admin/CompareGstFiles/CompareGSTCSV.cshtml");
						}
					}
					catch (Exception ex)
					{
						ViewBag.ErrorMessage = ex.Message;
						return View("/Views/Admin/CompareGstFiles/CompareGSTCSV.cshtml");
					}
				}
			}
			else
			{
				ViewBag.ErrorMessage = "Invalid file format.";
				return View("/Views/Admin/CompareGstFiles/CompareGSTCSV.cshtml");
			}

			try
			{
				// Save data to "GSTR2 DATA" table with data and ticket id
				await _gSTR2DataBusiness.SaveGSTR2DataAsync(dataTable, ticketnumber , ClientGSTIN);
			}
			catch (Exception ex)
			{
				ViewBag.ErrorMessage = $"{ex.Message} ";
				return View("/Views/Admin/CompareGstFiles/CompareGSTCSV.cshtml");
			}

			// Comparison logic (unchanged)
			DataTable gstr2Data = await _gSTR2DataBusiness.GetPortalDataBasedOnTicketAsync(ticketnumber , ClientGSTIN);
			//_logger.LogInformation("GSTR2 Columns: " + string.Join(", ", gstr2Data.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
			//_logger.LogInformation("Total GSTR2 Data Rows: " + gstr2Data.Rows.Count);

			DataTable purchaseData = await _purchaseDataBusiness.GetUserDataBasedOnTicketAsync(ticketnumber , ClientGSTIN);
			//_logger.LogInformation("purchaseData " + string.Join(", ", purchaseData.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
			//_logger.LogInformation("Total Purchase Data Rows: " + purchaseData.Rows.Count);

			await _compareGstBusiness.SaveCompareDataAsync(purchaseData, gstr2Data);
			await _purchaseTicketBusiness.UpdatePurchaseTicketAsync(ticketnumber, ClientGSTIN, AdminFileName);


			var data = await _compareGstBusiness.GetComparedDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
			ViewBag.ReportDataList = data;
			//_logger.LogInformation("Total Compared Data Rows: " + data.Count);

			// store number and sum in a model and save it in 
			ViewBag.Summary = GenerateSummary(data);
			var summary = ViewBag.Summary;
			//_logger.LogInformation($"invoice 1 count {summary.catagory1InvoiceNumber},portal 1 count {summary.catagory1PortalNumber}");

			return View("/Views/Admin/CompareGstFiles/Compare.cshtml");
		}

		private DataTable ReadCsvFile(Stream stream, string gstin)
		{
			var portalColumns = _configuration["Portal"].Split(',').Select(x => x.Trim()).ToList();

			DataTable dt = new DataTable();
			int lineNumber = 0;
			// usergstin , invno, invoice_date, suppliergstin,suppliername, returnperiod
			List<int> columnMismatchRows = new List<int>();

			List<int> userGstinInvalidRows = new List<int>();
			List<int> invoiceNumberInvalidRows = new List<int>();
			List<int> invoiceDateInvalidRows = new List<int>();
			List<int> supplierGstinInvalidRows = new List<int>();
			List<int> supplierNameInvalidRows = new List<int>();
			List<int> returnPeriodInvalidRows = new List<int>();

			using (var reader = new StreamReader(stream))
			using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(reader))
			{
				parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
				parser.SetDelimiters(",");
				parser.HasFieldsEnclosedInQuotes = true;

				// Read headers
				string[] headers = parser.ReadFields();
				lineNumber++;
				foreach (string header in headers)
				{
					dt.Columns.Add(header.Trim(), typeof(string));
				}
				//User GSTIN	GstRegType	invoiceno	invoice_date	SupplierGSTIN	SupplierName	IsRcmApplied	InvoiceValue	ItemTaxableValue	GstRate	IGSTAmount	CGSTAmount	SGSTAmount	CESS	IsReturnFiled	ReturnPeriod

				// usergstin , invno, invoice_date, suppliergstin,suppliername, returnperiod

				// Index mapping and validation
				int userGstinIndex = FindportalColumnIndex(headers, portalColumns[0]);
				int invoiceNumberColumnIndex = FindportalColumnIndex(headers, portalColumns[2]);
				int invoiceDateColumnIndex = FindportalColumnIndex(headers, portalColumns[3]);
				int supplierGstinIndex = FindportalColumnIndex(headers, portalColumns[4]);
				int supplierNameIndex = FindportalColumnIndex(headers, portalColumns[5]);
				int returnPeriodIndex = FindportalColumnIndex(headers, portalColumns[14]);

				// Read data rows
				while (!parser.EndOfData)
				{
					string[] fields = parser.ReadFields();
					lineNumber++;

					if (fields.Length != dt.Columns.Count)
					{
						columnMismatchRows.Add(lineNumber);
						continue;
					}
					// User GSTIN Validation
					string userGstinValue = fields[userGstinIndex];
					if (string.IsNullOrWhiteSpace(userGstinValue) || userGstinValue.Length != 15 || userGstinValue != gstin)
					{
						userGstinInvalidRows.Add(lineNumber);
						continue;
					}
					// Invoice Number Validation
					string invoiceNumberValue = fields[invoiceNumberColumnIndex];
					if (string.IsNullOrWhiteSpace(invoiceNumberValue) || invoiceNumberValue.Length > 25)
					{
						invoiceNumberInvalidRows.Add(lineNumber);
						continue;
					}
					// Invoice Date Validation
					string invoiceDateStr = fields[invoiceDateColumnIndex]?.Trim();
					string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
					if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
					{
						fields[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						invoiceDateInvalidRows.Add(lineNumber);
						continue;
					}
					// Supplier GSTIN Validation
					string supplierGstinValue = fields[supplierGstinIndex];
					if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
					{
						supplierGstinInvalidRows.Add(lineNumber);
						continue;
					}
					//Supplier Name Validation
					string supplierNameValue = fields[supplierNameIndex];
					if (supplierNameValue.Length > 100)
					{
						supplierNameInvalidRows.Add(lineNumber);
						continue;
					}
					// Return Period Validation
					string returnPeriodValue = fields[returnPeriodIndex];
					//if (string.IsNullOrWhiteSpace(returnPeriodValue))
					//{
					//    returnPeriodInvalidRows.Add(lineNumber);
					//    continue;
					//}

					dt.Rows.Add(fields);
				}
			}
			// Handle mismatch errors
			if (columnMismatchRows.Count > 0 || userGstinInvalidRows.Count > 0 || invoiceNumberInvalidRows.Count > 0 || invoiceDateInvalidRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || supplierNameInvalidRows.Count > 0 || returnPeriodInvalidRows.Count > 0)
			{
				var error = "";
				if (columnMismatchRows.Count > 0)
					error += $"Column count mismatch at line(s): {string.Join(", ", columnMismatchRows)}. Expected column count: {dt.Columns.Count}.";
				if (userGstinInvalidRows.Count > 0)
					error += $"\n Invalid User GSTIN at line(s): {string.Join(", ", userGstinInvalidRows)}.";
				if (invoiceNumberInvalidRows.Count > 0)
					error += $"\n The Invoice Number field allows a maximum of 25 characters.Invalid Invoice Number at line(s) : {string.Join(", ", invoiceNumberInvalidRows)}.";
				if (invoiceDateInvalidRows.Count > 0)
					error += $"\n Invoice date formate mismatch at line(s): {string.Join(", ", invoiceDateInvalidRows)}. Expected Formate : dd-MM-yyyy hh:mm:ss tt .";
				if (supplierGstinInvalidRows.Count > 0)
					error += $"\n Invalid Supplier GSTIN at line(s): {string.Join(", ", supplierGstinInvalidRows)}.";
				if (supplierNameInvalidRows.Count > 0)
					error += $"\n The Supplier Name field allows a maximum of 100 characters.Invalid Supplier Name at line(s) : {string.Join(", ", supplierNameInvalidRows)}.";
				if (returnPeriodInvalidRows.Count > 0)
					error += $"\n Period is null at line(s): {string.Join(", ", returnPeriodInvalidRows)}.";

				throw new Exception("Failed to insert data: \n" + error);
			}
			return dt;
		}
		private DataTable ReadExcelFile(Stream stream, string sheetName, string gstin)
		{
			var portalColumns = _configuration["Portal"].Split(',').Select(x => x.Trim()).ToList();

			DataTable dt = new DataTable();
			int lineNumber = 0;
			// usergstin , invno, invoice_date, suppliergstin,suppliername, returnperiod
			List<int> columnMismatchRows = new List<int>();

			List<int> userGstinInvalidRows = new List<int>();
			List<int> invoiceNumberInvalidRows = new List<int>();
			List<int> invoiceDateInvalidRows = new List<int>();
			List<int> supplierGstinInvalidRows = new List<int>();
			List<int> supplierNameInvalidRows = new List<int>();
			List<int> returnPeriodInvalidRows = new List<int>();


			using (var workbook = new XLWorkbook(stream))
			{
				//var worksheet = workbook.Worksheets.First();
				var worksheet = workbook.Worksheet(sheetName);
				if (worksheet == null)
					throw new Exception($"Sheet '{sheetName}' not found.");
				var rows = worksheet.RowsUsed().ToList();

				if (rows.Count == 0)
					throw new Exception("Excel file is empty.");

				// Read header
				var headerRow = rows[0];
				var headers = headerRow.Cells().Select(c => c.GetString().Trim()).ToArray();

				foreach (var header in headers)
					dt.Columns.Add(header, typeof(string));

				// Index mapping and validation
				int userGstinIndex = FindportalColumnIndex(headers, portalColumns[0]);
				int invoiceNumberColumnIndex = FindportalColumnIndex(headers, portalColumns[2]);
				int invoiceDateColumnIndex = FindportalColumnIndex(headers, portalColumns[3]);
				int supplierGstinIndex = FindportalColumnIndex(headers, portalColumns[4]);
				int supplierNameIndex = FindportalColumnIndex(headers, portalColumns[5]);
				int returnPeriodIndex = FindportalColumnIndex(headers, portalColumns[14]);

				// Read data rows
				for (int i = 1; i < rows.Count; i++)
				{

					var row = rows[i];
					var cells = row.Cells().Select((c, index) =>
					{
						if (index == invoiceDateColumnIndex) // Replace with the correct index for invoice_date
						{
							if (DateTime.TryParse(c.GetValue<string>(), out DateTime parsedDate))
							{
								return parsedDate.ToString("dd-MM-yyyy hh:mm:ss tt"); // or "yyyy - MM - dd" based on your DB requirement
							}
						}
						return c.GetFormattedString()?.Trim() ?? "";
					}).ToArray();

					int excelLineNumber = i + 1; // Line number as seen in Excel (header is line 1)
												 // Column count check
					if (cells.Length != dt.Columns.Count)
					{
						columnMismatchRows.Add(excelLineNumber);
						continue;
					}
					// User GSTIN Validation
					string userGstinValue = cells[userGstinIndex];
					if (string.IsNullOrWhiteSpace(userGstinValue) || userGstinValue.Length != 15 || userGstinValue != gstin)
					{
						userGstinInvalidRows.Add(excelLineNumber);
						continue;
					}
					// Invoice Number Validation
					string invoiceNumberValue = cells[invoiceNumberColumnIndex];
					if (string.IsNullOrWhiteSpace(invoiceNumberValue) || invoiceNumberValue.Length > 25)
					{
						invoiceNumberInvalidRows.Add(excelLineNumber);
						continue;
					}
					// Invoice Date Validation
					string invoiceDateStr = cells[invoiceDateColumnIndex]?.Trim();
					string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
					if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
					{
						cells[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						invoiceDateInvalidRows.Add(excelLineNumber);
						continue;
					}
					//Supplier GSTIN Validation
					string supplierGstinValue = cells[supplierGstinIndex];
					if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
					{
						supplierGstinInvalidRows.Add(excelLineNumber);
						continue;
					}
					//Supplier Name Validation
					string supplierNameValue = cells[supplierNameIndex];
					if (supplierNameValue.Length > 100)
					{
						supplierNameInvalidRows.Add(excelLineNumber);
						continue;
					}
					// Return Period Validation
					string returnPeriodValue = cells[returnPeriodIndex];
					//if (string.IsNullOrWhiteSpace(returnPeriodValue))
					//{
					//    returnPeriodInvalidRows.Add(excelLineNumber);
					//    continue;
					//}

					// All validations passed - Add to DataTable
					dt.Rows.Add(cells);
				}
			}
			// Handle mismatch errors
			if (columnMismatchRows.Count > 0 || userGstinInvalidRows.Count > 0 || invoiceNumberInvalidRows.Count > 0 || invoiceDateInvalidRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || supplierNameInvalidRows.Count > 0 || returnPeriodInvalidRows.Count > 0)
			{
				var error = "";
				if (columnMismatchRows.Count > 0)
					error += $"Column count mismatch at line(s): {string.Join(", ", columnMismatchRows)}. Expected column count: {dt.Columns.Count}.";
				if (userGstinInvalidRows.Count > 0)
					error += $"\n Invalid User GSTIN at line(s): {string.Join(", ", userGstinInvalidRows)}.";
				if (invoiceNumberInvalidRows.Count > 0)
					error += $"\n The Invoice Number field allows a maximum of 25 characters.Invalid Invoice Number at line(s) : {string.Join(", ", invoiceNumberInvalidRows)}.";
				if (invoiceDateInvalidRows.Count > 0)
					error += $"\n Invoice date formate mismatch at line(s): {string.Join(", ", invoiceDateInvalidRows)}. Expected Formate : dd-MM-yyyy hh:mm:ss tt .";
				if (supplierGstinInvalidRows.Count > 0)
					error += $"\n Invalid Supplier GSTIN at line(s): {string.Join(", ", supplierGstinInvalidRows)}.";
				if (supplierNameInvalidRows.Count > 0)
					error += $"\n The Supplier Name field allows a maximum of 100 characters.Invalid Supplier Name at line(s) : {string.Join(", ", supplierNameInvalidRows)}.";
				if (returnPeriodInvalidRows.Count > 0)
					error += $"\n Period is null at line(s): {string.Join(", ", returnPeriodInvalidRows)}.";

				throw new Exception("Failed to insert data: \n" + error);
			}
			return dt;

		}
		private DataTable ReadXLSFile(Stream stream, string sheetName, string gstin)
		{
			var portalColumns = _configuration["Portal"].Split(',').Select(x => x.Trim()).ToList();

			DataTable dt = new DataTable();
			int lineNumber = 0;
			// usergstin , invno, invoice_date, suppliergstin,suppliername, returnperiod
			List<int> columnMismatchRows = new List<int>();

			List<int> userGstinInvalidRows = new List<int>();
			List<int> invoiceNumberInvalidRows = new List<int>();
			List<int> invoiceDateInvalidRows = new List<int>();
			List<int> supplierGstinInvalidRows = new List<int>();
			List<int> supplierNameInvalidRows = new List<int>();
			List<int> returnPeriodInvalidRows = new List<int>();

			// Load the .xls workbook using NPOI
			HSSFWorkbook workbook = new HSSFWorkbook(stream);
			ISheet sheet = workbook.GetSheet(sheetName) ?? throw new Exception($"Sheet '{sheetName}' not found.");

			if (sheet.PhysicalNumberOfRows == 0)
				throw new Exception("Excel file is empty.");

			// Read header
			IRow headerRow = sheet.GetRow(0);
			int columnCount = headerRow.LastCellNum;
			string[] headers = new string[columnCount];
			for (int i = 0; i < columnCount; i++)
			{
				headers[i] = headerRow.GetCell(i)?.ToString().Trim() ?? $"Column{i}";
				dt.Columns.Add(headers[i], typeof(string));
			}
			// Index mapping and validation
			int userGstinIndex = FindportalColumnIndex(headers, portalColumns[0]);
			int invoiceNumberColumnIndex = FindportalColumnIndex(headers, portalColumns[2]);
			int invoiceDateColumnIndex = FindportalColumnIndex(headers, portalColumns[3]);
			int supplierGstinIndex = FindportalColumnIndex(headers, portalColumns[4]);
			int supplierNameIndex = FindportalColumnIndex(headers, portalColumns[5]);
			int returnPeriodIndex = FindportalColumnIndex(headers, portalColumns[14]);


			// Process rows
			for (int i = 1; i <= sheet.LastRowNum; i++)
			{
				IRow row = sheet.GetRow(i);
				if (row == null) continue;

				string[] cells = new string[columnCount];
				for (int j = 0; j < columnCount; j++)
				{
					ICell cell = row.GetCell(j);
					if (cell == null)
					{
						cells[j] = string.Empty;
						continue;
					}

					if (j == invoiceDateColumnIndex && cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
					{
						DateTime DT = (DateTime)cell.DateCellValue;
						cells[j] = DT.ToString("dd-MM-yyyy hh:mm:ss tt"); // Match your allowed format
					}
					else
					{
						cells[j] = cell.ToString().Trim();
					}
				}
				int excelLineNumber = i + 1; // Line number as seen in Excel (header is line 1)
											 // Column count check
				if (cells.Length != dt.Columns.Count)
				{
					columnMismatchRows.Add(excelLineNumber);
					continue;
				}
				// User GSTIN Validation
				string userGstinValue = cells[userGstinIndex];
				if (string.IsNullOrWhiteSpace(userGstinValue) || userGstinValue.Length != 15 || userGstinValue != gstin)
				{
					userGstinInvalidRows.Add(excelLineNumber);
					continue;
				}
				// Invoice Number Validation
				string invoiceNumberValue = cells[invoiceNumberColumnIndex];
				if (string.IsNullOrWhiteSpace(invoiceNumberValue) || invoiceNumberValue.Length > 25)
				{
					invoiceNumberInvalidRows.Add(excelLineNumber);
					continue;
				}
				// Invoice Date Validation
				string invoiceDateStr = cells[invoiceDateColumnIndex]?.Trim();
				string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
				if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
				{
					cells[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
				}
				else
				{
					invoiceDateInvalidRows.Add(excelLineNumber);
					continue;
				}
				//Supplier GSTIN Validation
				string supplierGstinValue = cells[supplierGstinIndex];
				if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
				{
					supplierGstinInvalidRows.Add(excelLineNumber);
					continue;
				}
				//Supplier Name Validation
				string supplierNameValue = cells[supplierNameIndex];
				if (supplierNameValue.Length > 100)
				{
					supplierNameInvalidRows.Add(excelLineNumber);
					continue;
				}
				// Return Period Validation
				string returnPeriodValue = cells[returnPeriodIndex];
				//if (string.IsNullOrWhiteSpace(returnPeriodValue))
				//{
				//    returnPeriodInvalidRows.Add(excelLineNumber);
				//    continue;
				//}

				// All validations passed - Add to DataTable
				dt.Rows.Add(cells);
			}
			// Handle mismatch errors
			if (columnMismatchRows.Count > 0 || userGstinInvalidRows.Count > 0 || invoiceNumberInvalidRows.Count > 0 || invoiceDateInvalidRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || supplierNameInvalidRows.Count > 0 || returnPeriodInvalidRows.Count > 0)
			{
				var error = "";
				if (columnMismatchRows.Count > 0)
					error += $"Column count mismatch at line(s): {string.Join(", ", columnMismatchRows)}. Expected column count: {dt.Columns.Count}.";
				if (userGstinInvalidRows.Count > 0)
					error += $"\n Invalid User GSTIN at line(s): {string.Join(", ", userGstinInvalidRows)}.";
				if (invoiceNumberInvalidRows.Count > 0)
					error += $"\n The Invoice Number field allows a maximum of 25 characters.Invalid Invoice Number at line(s) : {string.Join(", ", invoiceNumberInvalidRows)}.";
				if (invoiceDateInvalidRows.Count > 0)
					error += $"\n Invoice date formate mismatch at line(s): {string.Join(", ", invoiceDateInvalidRows)}. Expected Formate : dd-MM-yyyy hh:mm:ss tt .";
				if (supplierGstinInvalidRows.Count > 0)
					error += $"\n Invalid Supplier GSTIN at line(s): {string.Join(", ", supplierGstinInvalidRows)}.";
				if (supplierNameInvalidRows.Count > 0)
					error += $"\n The Supplier Name field allows a maximum of 100 characters.Invalid Supplier Name at line(s) : {string.Join(", ", supplierNameInvalidRows)}.";
				if (returnPeriodInvalidRows.Count > 0)
					error += $"\n Period is null at line(s): {string.Join(", ", returnPeriodInvalidRows)}.";

				throw new Exception("Failed to insert data: \n" + error);
			}
			return dt;
		}
		private bool ValidateColumnNames(DataTable dataTable)
		{
			var _portal = _configuration["Portal"];
			var _portalColumns = _portal.Split(',').Select(x => x.Trim()).ToList();
			var uploadedColumns = dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName.ToLower()).ToList();
			return _portalColumns.All(col => uploadedColumns.Contains(col.ToLower()));
		}
		private int FindportalColumnIndex(string[] headers, string expectedName)
		{
			int index = Array.FindIndex(headers, h => h.Equals(expectedName, StringComparison.OrdinalIgnoreCase));
			if (index == -1)
				throw new Exception($"The Portal file does not contain a '{expectedName}' column.");
			return index;
		}

		#endregion

		#region getUnMatchedData
		public async Task<IActionResult> getUnMatchedData(string requestNo, string ClientGstin, string matchType)
		{
			ViewBag.Messages = "Admin";
			DataTable UnmatchedData = await _compareGstBusiness.GetUnMatchedData(requestNo, ClientGstin);
			DataTable invoiceTable = UnmatchedData.Clone(); // Clones the structure
			DataTable portalTable = UnmatchedData.Clone();

			//Console.WriteLine("getUnMatchedData UnmatchedData: " + UnmatchedData.Rows.Count);
		   // Console.WriteLine("Match type : " + matchType);

			string matchType7 = _configuration["MatchTypes:7"];
			string matchType8 = _configuration["MatchTypes:8"];
			string matchType6 = _configuration["MatchTypes:6"];

			if (matchType == matchType7 || matchType == matchType8)
			{
				//Console.WriteLine("getUnMatchedData matchType7 or matchType8");
				foreach (DataRow row in UnmatchedData.Rows)
				{
					if (row["DataSource"].ToString().Trim().Equals("Invoice", StringComparison.OrdinalIgnoreCase)
						&& row["MatchType"].ToString().Trim() == matchType7)
					{
						invoiceTable.ImportRow(row);
					}
					else if (row["DataSource"].ToString().Trim().Equals("Portal", StringComparison.OrdinalIgnoreCase)
						&& row["MatchType"].ToString().Trim() == matchType8)
					{
						portalTable.ImportRow(row);
					}
				}
			}
			else
			{
				//Console.WriteLine("Match type 6");
				foreach (DataRow row in UnmatchedData.Rows)
				{
					if (row["DataSource"].ToString().Trim().Equals("invoice", StringComparison.OrdinalIgnoreCase)
						&& row["MatchType"].ToString() == matchType6)
					{
						invoiceTable.ImportRow(row);
					}
					else if (row["DataSource"].ToString().Trim().Equals("portal", StringComparison.OrdinalIgnoreCase)
						&& row["MatchType"].ToString() == matchType6)
					{
						portalTable.ImportRow(row);
					}
				}
			}


			//Console.WriteLine("getUnMatchedData Invoice Table: " + invoiceTable.Rows.Count);
		   // Console.WriteLine("Invoice Table Headers: " + string.Join(", ", invoiceTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName)));
		   // foreach (DataRow row in invoiceTable.Rows)
			//    Console.WriteLine(string.Join(", ", row.ItemArray));


			//Console.WriteLine("getUnMatchedData Portal Table: " + portalTable.Rows.Count);

			var model = new EditableTablesViewModel
			{
				InvoiceTable = ConvertToDictionaryList(invoiceTable),
				PortalTable = ConvertToDictionaryList(portalTable)
			};

			ViewBag.requestNo = requestNo;
			ViewBag.ClientGstin = ClientGstin;
			ViewBag.matchType = matchType;
			return View("~/Views/Admin/CompareGstFiles/ModifyData.cshtml", model);
		}

		private List<Dictionary<string, object>> ConvertToDictionaryList(DataTable dt)
		{
			var list = new List<Dictionary<string, object>>();
			foreach (DataRow row in dt.Rows)
			{
				var dict = new Dictionary<string, object>();
				foreach (DataColumn col in dt.Columns)
				{
					dict[col.ColumnName] = row[col];
				}
				list.Add(dict);
			}
			return list;
		}

		[HttpPost]
		public async Task<IActionResult> EditCompare(string RequestNo, string ClientGSTIN, string InvoiceTable, string PortalTable)
		{
			ViewBag.Messages = "Admin";
			try
			{
				// Deserialize JSON to DataTable
				DataTable invoiceData = JsonConvert.DeserializeObject<DataTable>(InvoiceTable);
				DataTable portalData = JsonConvert.DeserializeObject<DataTable>(PortalTable);

				//Console.WriteLine("EditCompare invoice count : " + invoiceData.Rows.Count);
			   // Console.WriteLine("Invoice Table Headers: " + string.Join(", ", invoiceData.Columns.Cast<DataColumn>().Select(col => col.ColumnName)));
			   // foreach (DataRow row in invoiceData.Rows)
				   // Console.WriteLine(string.Join(", ", row.ItemArray));

				//Console.WriteLine("EditCompare portal count : " + portalData.Rows.Count);
				// Get ticket info
				var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(RequestNo, ClientGSTIN);
				var AdminFileName = Ticket.AdminFileName;


				//Console.ReadKey();

				// Save the modified data

				await _ModifiedDataBusiness.SaveModifiedData(invoiceData, RequestNo, ClientGSTIN);

				await _ModifiedDataBusiness.SaveModifiedData(portalData, RequestNo, ClientGSTIN);



				DataTable invoiceData1 = await _ModifiedDataBusiness.GetModifiedInvoiceDataBasedOnTicketAsync(RequestNo, ClientGSTIN);
				//Console.WriteLine("EditCompare invoice1 count : " + invoiceData1.Rows.Count);
				//Console.WriteLine("Invoice1 Table Headers: " + string.Join(", ", invoiceData1.Columns.Cast<DataColumn>().Select(col => col.ColumnName)));
				//foreach (DataRow row in invoiceData1.Rows)
				//    Console.WriteLine(string.Join(", ", row.ItemArray));

				// Get modified portal data
				DataTable gstr2Data = await _ModifiedDataBusiness.GetModifiedPortalDataBasedOnTicketAsync(RequestNo, ClientGSTIN);

				// Optionally save compared data
				await _compareGstBusiness.SaveCompareDataAsync(invoiceData1, gstr2Data);

				// Update ticket info
				await _purchaseTicketBusiness.UpdatePurchaseTicketAsync(RequestNo, ClientGSTIN, AdminFileName);

				// Get compared data
				var data = await _compareGstBusiness.GetComparedDataBasedOnTicketAsync(RequestNo, ClientGSTIN);

				ViewBag.Ticket = Ticket;
				ViewBag.ReportDataList = data;
				ViewBag.Summary = GenerateSummary(data);

				return View("/Views/Admin/CompareGstFiles/Compare.cshtml");
			}
			catch (Exception ex)
			{
				// Log error
				_logger.LogError(ex, "Error in EditCompare");
				return StatusCode(500, "An error occurred while comparing data.");
			}
		}

        #endregion

        #region PurchaseRegisterAPI-Master
        public IActionResult OpenRequestsAPI_Master()
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2"; // For active button styling
            return RedirectToAction("OpenTaskAPI_Master", "Admin");
        }

        public async Task<IActionResult> OpenTaskAPI_Master(DateTime? fromdate, DateTime? todate)
        {

            // Default to 2 days ago and today if no values are provided
            DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
            DateTime toDateTime = todate ?? DateTime.Now; // Default to today

            //await TjCaptions("OpenTasks"); // Load captions for OpenTask page
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                return View("~/Views/Admin/OpenTasks/OpenTaskAPI_Master.cshtml");
            }

            //var email = MySession.Current.Email; // Get the email from session
            //var clients = await _userBusiness.GetAdminClients(email);
            //string[] gstinArray = clients.Select(c => c.ClientGSTIN).ToArray();
            //var Tickets = await _purchaseTicketBusiness.GetClientsOpenTicketsStatusAsync(gstinArray, fromDateTime, toDateTime);

            var Tickets = await _purchaseTicketBusiness.GetOpenTicketsStatusAsync(MySession.Current.gstin, fromDateTime, toDateTime); //(ViewBag.UserGSTIN); // Fetch pending tickets based on user_gstin

            // HttpContext.Session.SetString("TicketsList", JsonConvert.SerializeObject(Tickets));
            ViewBag.Tickets = Tickets;

            return View("~/Views/Admin/OpenTasks/OpenTaskAPI_Master.cshtml"); // Admin dashboard
        }

        public async Task<IActionResult> CompareGSTAPI_Master(string ticketnumber, string ClientGSTIN, string fromDate, string toDate)
        {
            // _logger.LogInformation($"Ticket Number: {ticketnumber}");
            //based on this ticket number fetch data from Purchase Ticket table and store
            var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            ViewBag.Ticket = Ticket;
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            ViewBag.fromDate = fromDate;
            ViewBag.toDate = toDate;
            return View("~/Views/Admin/CompareGstFiles/CompareGSTAPI_Master.cshtml");
        }

        [HttpPost]
        public async Task<IActionResult> CompareAPI_Master(string ticketnumber, string ClientGSTIN)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            string sessionId = HttpContext.Session.Id;
            var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            ViewBag.Ticket = Ticket;

            string UserName = "";
            try
            {
                var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);
                UserName = UserAPIData.GstPortalUsername;
            }
            catch
            {
                string errorMessage = "Error fetching user API data. Please update user API data.";
                return Json(new { failure = true, message = errorMessage });
            }

            string Email = _configuration["PRMasterAPI:email"];

            try
            {
                var authTokenData = await _gSTR2DataBusiness.GetTokenDataAsync(ClientGSTIN);

                bool isTokenExpired = authTokenData == null ||
                                      string.IsNullOrEmpty(authTokenData.AuthToken) ||
                                      (authTokenData.AuthTokenCreatedDatetime?.AddSeconds(Convert.ToDouble(authTokenData.Expiry)) < DateTime.Now);

                if (!isTokenExpired)
                {
                    return Json(new { success = true, message = "Token is valid. Continue." });
                }

                #region Request for authentation  1st Api call

                // Parameters
                string Parameters = $"email={Email}";

                // Add required headers
                string userName = UserName;
                string stateCd = ClientGSTIN.Substring(0, 2);
                string ipAddress = _configuration["PRMasterAPI:ipAddress"];
                string clientId = _configuration["PRMasterAPI:ClientId"];
                string clientSecret = _configuration["PRMasterAPI:ClientSecret"];

                var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Add("gst_username", userName);
                httpClient.DefaultRequestHeaders.Add("state_cd", stateCd);
                httpClient.DefaultRequestHeaders.Add("ip_address", ipAddress);
                httpClient.DefaultRequestHeaders.Add("client_id", clientId);
                httpClient.DefaultRequestHeaders.Add("client_secret", clientSecret);

                // Body
                var RequestBody = new { };

                // Url
                string baseUrl = _configuration["PRMasterAPI:BaseUrl"] ?? throw new InvalidOperationException("BaseUrl for PRMasterAPI is not configured in ApiSettings");
                string otpRequestEndpoint = _configuration["PRMasterAPI:OtpRequest"] ?? throw new InvalidOperationException("OtpRequest for PRMasterAPI is not configured in ApiSettings");
                string apiUrl1 = $"{baseUrl}{otpRequestEndpoint}?{Parameters}";


                // API call
                var response = await httpClient.GetAsync(apiUrl1);
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);

                // save API data to database
                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                var headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketnumber;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl1;
                    ApiData.RequestParameters = Parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _gSTR2DataBusiness.savePortalAPIsData(ApiData);


                // Validate response 
                //Result 1  {
                //	"status_cd": "1",
                //	"status_desc": "user name exists",
                //	"header": {
                //		"gst_username": "MH_NT2.1642",
                //		"state_cd": "27",
                //		"ip_address": "14.98.237.54",
                //		"client_id": "GSTS36731818-d2d3-4036-816c-d3e559eb1dde",
                //		"client_secret": "GSTSa012637b-1daa-43be-aa1f-85d83ce132fb",
                //		"cache-control": "no-cache",
                //		"postman-token": "446ef596-b536-4e9a-8f71-887af9324664",
                //		"txn": "5dcd07ed2a714698b9e74ccea93bf775"
                //	}
                //}
                string txn = null;

                if (result["header"] != null && result["header"]["txn"] != null)
                {
                    txn = result["header"]["txn"].ToString();
                    return Json(new { success = false, askForOtp = true, txn = txn });
                }
                if (result["status_cd"].ToString() != "1" || txn == null)
                {
                    string errorMessage = $"{result["error"]["message"]}";
                    return Json(new
                    {
                        failure = true,
                        message = $"Purchase Register OTP Request API Call - Failed due to :{errorMessage}" //Msg from result 
                    });
                }

                return Json(new { failure = true, message = "Purchase Register OTP Request API Call - Failed ." });

                #endregion
            }
            catch (Exception ex)
            {
                string errorMessage = $"Error : {ex.Message}";
                return Json(new { failure = true, message = errorMessage });
            }
        }

        [HttpPost]
        public async Task<IActionResult> SubmitOtpAndContinue_Master(string ClientGSTIN, string ticketNo, string otp, string txn)
        {
            string sessionId = HttpContext.Session.Id;
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);

            string UserName = UserAPIData.GstPortalUsername;
            string Email = _configuration["PRMasterAPI:email"];

            try
            {
                #region 2nd Api call

                // Parameters
                string Parameters = $"email={Email}&otp={otp}";

                // Add required headers
                string stateCd = ClientGSTIN.Substring(0, 2);
                string ipAddress = _configuration["PRMasterAPI:ipAddress"];
                string clientId = _configuration["PRMasterAPI:ClientId"];
                string clientSecret = _configuration["PRMasterAPI:ClientSecret"];

                var httpClient = _httpClientFactory.CreateClient();

                httpClient.DefaultRequestHeaders.Add("gst_username", UserName);
                httpClient.DefaultRequestHeaders.Add("state_cd", stateCd);
                httpClient.DefaultRequestHeaders.Add("ip_address", ipAddress);
                httpClient.DefaultRequestHeaders.Add("txn", txn);
                httpClient.DefaultRequestHeaders.Add("client_id", clientId);
                httpClient.DefaultRequestHeaders.Add("client_secret", clientSecret);

                // Body
                var RequestBody = new { };

                // Url
                string baseUrl = _configuration["PRMasterAPI:BaseUrl"] ?? throw new InvalidOperationException("BaseUrl for PRMasterAPI is not configured in ApiSettings");
                string authTokenEndpoint = _configuration["PRMasterAPI:AuthToken"] ?? throw new InvalidOperationException("AuthToken for PRMasterAPI is not configured in ApiSettings");
                string apiUrl = $"{baseUrl}{authTokenEndpoint}?{Parameters}";

                // API call
                var response = await httpClient.GetAsync(apiUrl);
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);

                // save API data to database
                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                var headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketNo;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = Parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _gSTR2DataBusiness.savePortalAPIsData(ApiData);

                // Validate response 

                Console.WriteLine($"Response 2 : {responseData}");
                Console.WriteLine($"Result 2 : {result}");
                //Console.ReadKey();

                //Result 2 {
                //	"status_cd": "1",
                //	"status_desc": "If authentication succeeds",
                //	"header": {
                //		"gst_username": "MH_NT2.1642",
                //		"state_cd": "27",
                //		"ip_address": "14.98.237.54",
                //		"client_id": "GSTS36731818-d2d3-4036-816c-d3e559eb1dde",
                //		"client_secret": "GSTSa012637b-1daa-43be-aa1f-85d83ce132fb",
                //		"txn": "5dcd07ed2a714698b9e74ccea93bf775",
                //		"cache-control": "no-cache",
                //		"postman-token": "2dbb2361-fa6d-40f4-b8df-925b17443db5"
                //	}
                //}
                if (result["status_cd"].ToString() == "1")
                {
                    await _gSTR2DataBusiness.saveTokenData(new GSTR2TokenDataModel
                    {
                        ClientGstin = ClientGSTIN,
                        RequestNumber = ticketNo,
                        UserName = UserName,
                        XAppKey = txn,
                        OTP = otp,
                        AuthToken = "",
                        Expiry = "",
                        SEK = ""
                    });
                    return Json(new { success = true, message = "OTP Verified. Token saved." });
                }

                //Result 2 : {                            // Invalid otp response 
                //    "status_cd": "0",
                //    "status_desc": "If authentication fails",
                //    "error": {
                //        "message": "Invalid Session",
                //        "error_cd": "AUTH4033"
                //    },
                //    "header": {
                //        "gst_username": "MH_NT2.1642",
                //        "state_cd": "27",
                //        "ip_address": "14.98.237.54",
                //        "client_id": "GSTS36731818-d2d3-4036-816c-d3e559eb1dde",
                //        "client_secret": "GSTSa012637b-1daa-43be-aa1f-85d83ce132fb",
                //        "txn": "5dcd07ed2a714698b9e74ccea93bf775",
                //        "cache-control": "no-cache",
                //        "postman-token": "480c5a42-e634-482f-933c-23219ddf24b5"
                //    }
                //}
                if (result["status_cd"].ToString() == "0" && result["error"]["error_cd"].ToString() == "AUTH4033")
                {
                    return Json(new { success = false, askAgain = true, message = "Invalid OTP. Please enter again." });
                }

                if (result["status_cd"].ToString() != "1")
                {
                    string errorMessage = $"Purchase Register Auth Token API Call - Failed due to : {result["error"]["message"]}";
                    return Json(new { failure = true, message = errorMessage });
                }

                return Json(new { failure = true, message = "Purchase Register Auth Token API Call - Failed." });

                #endregion

            }
            catch (Exception ex)
            {
                string errorMessage = $"Error: {ex.Message}";
                return Json(new { failure = true, message = errorMessage });
            }
        }

        [HttpPost]
        public async Task<IActionResult> ContinueWith4thApi_Master(string ticketnumber, string ClientGSTIN)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            string sessionId = HttpContext.Session.Id;
            var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            ViewBag.Ticket = Ticket;

            string Txn_Period = Ticket.TxnPeriod;
            DateTime parsedDate;
            string period = "";
            // Parse the input string using the correct format
            if (DateTime.TryParseExact(Txn_Period, "MMM-yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
            {
                period = parsedDate.ToString("MMyyyy");
            }

            var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);

            string UserName = UserAPIData.GstPortalUsername;
            string Email = _configuration["PRMasterAPI:email"];

            var authTokenData = await _gSTR2DataBusiness.GetTokenDataAsync(ClientGSTIN);
            try
            {
                #region 4th Api call

                // Parameters
                string Parameters = $"email={Email}&gstin={ClientGSTIN}&rtnprd={period}";

                // Add required headers
                string stateCd = ClientGSTIN.Substring(0, 2);
                string ipAddress = _configuration["PRMasterAPI:ipAddress"];
                string clientId = _configuration["PRMasterAPI:ClientId"];
                string clientSecret = _configuration["PRMasterAPI:ClientSecret"];

                var httpClient = _httpClientFactory.CreateClient();

                httpClient.DefaultRequestHeaders.Add("gst_username", UserName);
                httpClient.DefaultRequestHeaders.Add("state_cd", stateCd);
                httpClient.DefaultRequestHeaders.Add("ip_address", ipAddress);
                httpClient.DefaultRequestHeaders.Add("txn", authTokenData.XAppKey);
                httpClient.DefaultRequestHeaders.Add("client_id", clientId);
                httpClient.DefaultRequestHeaders.Add("client_secret", clientSecret);

                // Body
                var RequestBody = new { };

                // Url
                string baseUrl = _configuration["PRMasterAPI:BaseUrl"] ?? throw new InvalidOperationException("BaseUrl for PRMasterAPI is not configured in ApiSettings");
                string getDataEndpoint = _configuration["PRMasterAPI:PortalData"] ?? throw new InvalidOperationException("PortalData endpoint for PRMasterAPI is not configured in ApiSettings");
                string apiUrl = $"{baseUrl}{getDataEndpoint}?{Parameters}";

                // API call
                var response = await httpClient.GetAsync(apiUrl); // ✅ Corrected: GET, not POST
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);

                // save API data to database

                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                var headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketnumber;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = Parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _gSTR2DataBusiness.savePortalAPIsData(ApiData);


                // Validate response 
                Console.WriteLine($"Response 4 : {responseData}");
                Console.WriteLine($"Result 4 : {result}");
                //Console.ReadKey();

                //

                //condition 

                //DataTable dataTable;
                //try
                //{
                //    dataTable = ConvertJsonToDataTable2(responseData, ClientGSTIN);
                //    await _gSTR2DataBusiness.SaveGSTR2DataAsync(dataTable, ticketnumber, ClientGSTIN);
                //    return Json(new { goToCompare = true });
                //}
                //catch (Exception ex)
                //{
                //    string errorMessage = $"Error: {ex.Message}";
                //    return Json(new { failure = true, message = errorMessage });
                //}


                //Result 3 : {
                //    "data": {
                //        "est": "30",
                //		  "token": "d889e2328ec949aabf9b2be90aabba7c"
                //    },
                //	"status_cd": "2",
                //	"status_desc": "GSTR request succeeds",
                //	"header": {
                //        "gst_username": "MH_NT2.1642",
                //		"state_cd": "27",
                //		"ip_address": "14.98.237.54",
                //		"client_id": "GSTS36731818-d2d3-4036-816c-d3e559eb1dde",
                //		"client_secret": "GSTSa012637b-1daa-43be-aa1f-85d83ce132fb",
                //		"txn": "5dcd07ed2a714698b9e74ccea93bf775",
                //		"cache-control": "no-cache",
                //		"postman-token": "cee1bc5e-8e6b-451f-b288-d561dc0781e1",
                //		"gstin": "27AAGCB1286Q2Z3"
                //    }
                //}
                string token = null;
                // check if "data" node exists
                if (result["data"] != null && result["data"]["token"] != null)
                {
                    token = result["data"]["token"].ToString();
                    return Json(new { success = true, token = token });
                }

                //Result 4 : {
                //			  "status_cd": "0",
                //			  "status_desc": "GSTR request failed",
                //			  "error": {
                //				"message": "No document found for the provided Inputs",
                //				"error_cd": "RET13509"
                //			  },

                if (result["status_cd"]?.ToString() != "1" && result["error"]["message"]?.ToString() == "No document found for the provided Inputs")
                {
                    return Json(new { goToCompare = true });
                }

                //Result 4 : {
                //              "status_cd": "0",
                //              "status_desc": "GSTR request failed",
                //              "error": {
                //                "message": "Either there is no record for GSTR 2B or GSTR 2B for current return period is not generated by system (due to non-filing GSTR 3B of last return period till GSTR 2B cut-off date of current return period), kindly generate 2B ondemand",
                //                "error_cd": "GTR2B-002"
                //              },
                //              "header": {
                //                "gst_username": "TN_NT2.152384",
                //                "state_cd": "33",
                //                "ip_address": "14.98.237.54",
                //                "txn": "166fdfca61b44615bfec1dbbf0ce7d5a",
                //                "client_id": "GSTS36731818-d2d3-4036-816c-d3e559eb1dde",
                //                "client_secret": "GSTSa012637b-1daa-43be-aa1f-85d83ce132fb",
                //                "traceparent": "00-22bf1beccafd26b1b492e06955348efb-07b28ecd6f193e24-00",
                //                "ret_period": "042025",
                //                "gstin": "33AAGCB1286Q2ZA"
                //              }
                //            }

                if (result["status_cd"]?.ToString() != "1" && result["error"]["message"].ToString().Contains("Either there is no record for GSTR 2B"))
                {
                    return Json(new { goToCompare = true });
                }

                if (result["status_cd"]?.ToString() != "1" || string.IsNullOrEmpty(token))
                {
                    string errorMessage = $"Purchase Register GET Invoices API Call - Failed due to  : {result["error"]["message"]}";
                    return Json(new { failure = true, message = errorMessage });
                }

                return Json(new { failure = true, message = "Purchase Register GET Invoices API Call - Failed." });

                #endregion
            }
            catch (Exception ex)
            {
                string errorMessage = $"Error: {ex.Message}";
                return Json(new { failure = true, message = errorMessage });
            }
        }

        public async Task<IActionResult> CompareDataTables_Master(string ticketnumber, string ClientGSTIN)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            ViewBag.Ticket = Ticket;

            // Comparison logic (unchanged)
            DataTable gstr2Data = await _gSTR2DataBusiness.GetPortalDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            //_logger.LogInformation("GSTR2 Columns: " + string.Join(", ", gstr2Data.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
            //_logger.LogInformation("Total GSTR2 Data Rows: " + gstr2Data.Rows.Count);

            DataTable purchaseData = await _purchaseDataBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            //_logger.LogInformation("purchaseData " + string.Join(", ", purchaseData.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
            //_logger.LogInformation("Total Purchase Data Rows: " + purchaseData.Rows.Count);

            await _compareGstBusiness.SaveCompareDataAsync(purchaseData, gstr2Data);

            string AdminFileName = "API_Master.csv";
            await _purchaseTicketBusiness.UpdatePurchaseTicketAsync(ticketnumber, ClientGSTIN, AdminFileName);

            var data = await _compareGstBusiness.GetComparedDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            ViewBag.ReportDataList = data;
            //_logger.LogInformation("Total Compared Data Rows: " + data.Count);

            // store number and sum in a model and save it in 
            ViewBag.Summary = GenerateSummary(data);
            var summary = ViewBag.Summary;
            //_logger.LogInformation($"invoice 1 count {summary.catagory1InvoiceNumber},portal 1 count {summary.catagory1PortalNumber}");

            return View("/Views/Admin/CompareGstFiles/Compare.cshtml");
        }

        private DataTable ConvertJsonToDataTable2(string json, string userGstin)
        {
            var table = new DataTable();

            // Define columns
            table.Columns.Add("User GSTIN");
            table.Columns.Add("GstRegType");
            table.Columns.Add("invoiceno");
            table.Columns.Add("invoice_date");
            table.Columns.Add("Supplier GSTIN");
            table.Columns.Add("SupplierName");
            table.Columns.Add("IsRcmApplied");
            table.Columns.Add("InvoiceValue", typeof(decimal));
            table.Columns.Add("ItemTaxableValue", typeof(decimal));
            table.Columns.Add("GstRate", typeof(decimal));
            table.Columns.Add("IGSTAmount", typeof(decimal));
            table.Columns.Add("CGSTAmount", typeof(decimal));
            table.Columns.Add("SGSTAmount", typeof(decimal));
            table.Columns.Add("CESS", typeof(decimal));
            table.Columns.Add("IsReturnFiled");
            table.Columns.Add("ReturnPeriod");

            JObject root = JObject.Parse(json);
            var retPeriod = root["header"]?["ret_period"]?.ToString();

            var b2bList = root["data"]?["b2b"] as JArray;
            if (b2bList == null) return table;

            foreach (var b2b in b2bList)
            {
                string cfs = b2b["cfs"]?.ToString(); // IsReturnFiled
                string ctin = b2b["ctin"]?.ToString();
                var invoices = b2b["inv"] as JArray;

                foreach (var inv in invoices)
                {
                    string inum = inv["inum"]?.ToString();
                    string idt = inv["idt"]?.ToString();
                    string inv_typ = inv["inv_typ"]?.ToString();
                    string rchrg = inv["rchrg"]?.ToString();
                    decimal val = inv["val"]?.ToObject<decimal>() ?? 0;

                    var items = inv["itms"] as JArray;
                    foreach (var item in items)
                    {
                        var details = item["itm_det"];
                        decimal txval = details["txval"]?.ToObject<decimal>() ?? 0;
                        decimal rt = details["rt"]?.ToObject<decimal>() ?? 0;
                        decimal iamt = details["iamt"]?.ToObject<decimal>() ?? 0;
                        decimal camt = details["camt"]?.ToObject<decimal>() ?? 0;
                        decimal samt = details["samt"]?.ToObject<decimal>() ?? 0;
                        decimal csamt = details["csamt"]?.ToObject<decimal>() ?? 0;

                        table.Rows.Add(
                           userGstin,         // User GSTIN
                           inv_typ,           // GstRegType
                           inum,              // invoiceno
                           idt,               // invoice_date
                           ctin,              // SupplierGSTIN
                           "",                // SupplierName (Not provided)
                           "",                // IsRcmApplied (Not provided)
                           val,               // InvoiceValue
                           txval,             // ItemTaxableValue
                           rt,                // GstRate
                           iamt,              // IGSTAmount
                           camt,              // CGSTAmount
                           samt,              // SGSTAmount
                           csamt,             // CESS
                           cfs,                // IsReturnFiled (Not provided)
                           retPeriod          // ReturnPeriod
                       );
                    }
                }
            }

            return table;
        }

        #endregion

        #region PurchaseRegisterAPI-IRIS
        public IActionResult OpenRequestsAPI_Iris()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act2"; // For active button styling
			return RedirectToAction("OpenTaskAPI_IRIS", "Admin");
		}

		public async Task<IActionResult> OpenTaskAPI_IRIS(DateTime? fromdate, DateTime? todate)
		{
			// Default to 2 days ago and today if no values are provided
			DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
			DateTime toDateTime = todate ?? DateTime.Now; // Default to today

			//await TjCaptions("OpenTasks"); // Load captions for OpenTask page
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act2"; // For active button styling
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                return View("~/Views/Admin/OpenTasks/OpenTaskAPI_IRIS.cshtml");
            }

            //var Tickets = await _purchaseTicketBusiness.GetAllOpenTicketsStatusAsync(fromDateTime, toDateTime);
            //         var email = MySession.Current.Email; // Get the email from session
            //var clients = await _userBusiness.GetAdminClients(email);
            //string[] gstinArray = clients.Select(c => c.ClientGSTIN).ToArray();
            //var Tickets = await _purchaseTicketBusiness.GetClientsOpenTicketsStatusAsync(gstinArray, fromDateTime, toDateTime);

            var Tickets = await _purchaseTicketBusiness.GetOpenTicketsStatusAsync(MySession.Current.gstin, fromDateTime, toDateTime); //(ViewBag.UserGSTIN); // Fetch pending tickets based on user_gstin

            // HttpContext.Session.SetString("TicketsList", JsonConvert.SerializeObject(Tickets));
            ViewBag.Tickets = Tickets;

			return View("~/Views/Admin/OpenTasks/OpenTaskAPI_IRIS.cshtml"); // Admin dashboard

			//based on his clients Gstin Fetch tickets and show in opentickets page
			//step -1
			// based on his id get the client gstin list from Admin_client_link table
			//let say this Admin with back@end.com mail id have 3 clients their gstin is {1,2,3} 
			//string[] clients = { "1", "2", "3" }; // create and call function here
			//step -2
			//Based on this clients Gstin fetch Pending the tickets from Purchase_Ticket table
			//var Tickets = await _purchaseTicketBusiness.GetClientTicketsStatusAsync(clients);
		}

		public async Task<IActionResult> CompareGSTAPI_IRIS(string ticketnumber, string ClientGSTIN, string fromDate, string toDate)
		{
			//Console.WriteLine($"Ticket Number: {ticketnumber}");
            //Console.WriteLine($"ClientGSTIN Number: {ClientGSTIN}");
			//based on this ticket number fetch data from Purchase Ticket table and store
			var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);

			ViewBag.Ticket = Ticket;
			ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            ViewBag.fromDate = fromDate;
            ViewBag.toDate = toDate;

            return View("~/Views/Admin/CompareGstFiles/CompareGSTAPI_IRIS.cshtml");
		}

        [HttpPost]
        public async Task<IActionResult> CompareAPI_IRIS(string ticketnumber, string ClientGSTIN)
		{
				
            string sessionId = HttpContext.Session.Id;
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act2";
            var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
			ViewBag.Ticket = Ticket;

			//return Json(new
			//{
			//	success = false,
			//	askForOtp = true
			//});

			string Txn_Period = Ticket.TxnPeriod;
			string userName = "";

            try 
			{
                var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);
                userName = UserAPIData.GstPortalUsername;
            }
			catch
			{
				string errorMessage = "Error fetching user API data. Please update user API data.";
				return Json(new { message = errorMessage });
			}
            
            string clientid = _configuration["PurchaseRegisterIRISAPI:ClientId"];
            string clientSecret = _configuration["PurchaseRegisterIRISAPI:ClientSecret"];
            string statecd = ClientGSTIN.Substring(0, 2);  // Extracting state code from GSTIN
            string ipusr = _configuration["PurchaseRegisterIRISAPI:ipurs"];
            string txn = _configuration["PurchaseRegisterIRISAPI:txn"];

			string parameters;
			var httpClient = _httpClientFactory.CreateClient();
			string baseUrl , AuthUrl , apiUrl , headersJson;
         
			#region API calls to get Portal Data

            try
			{
                // Step 1: Get existing auth token from DB

                // get authtoken and authtokenCreateddatetime and expiry from Database  - gstin,ticketno
                var authTokenData = await _gSTR2DataBusiness.GetTokenDataAsync(ClientGSTIN);

				bool isTokenExpired = authTokenData == null ||
									  string.IsNullOrEmpty(authTokenData.AuthToken) ||
									  (authTokenData.AuthTokenCreatedDatetime?.AddSeconds(Convert.ToDouble(authTokenData.Expiry)) < DateTime.Now);

				//bool isTokenExpired = authTokenData == null ||
				//					  string.IsNullOrEmpty(authTokenData.AuthToken) ||
				//					  (authTokenData.AuthTokenCreatedDatetime?.AddMinutes(Convert.ToDouble(authTokenData.Expiry)) < DateTime.Now);

				//Console.WriteLine($"ticketnumber : {ticketnumber}");
				//Console.WriteLine($"ClientGSTIN : {ClientGSTIN}");

				//Console.WriteLine($"AuthTokenCreatedDatetime: {authTokenData.AuthTokenCreatedDatetime}");
				//Console.WriteLine($"Expiry: {authTokenData.Expiry}");
				//Console.WriteLine($"Auth token valid upto : {authTokenData.AuthTokenCreatedDatetime?.AddSeconds(Convert.ToDouble(authTokenData.Expiry))}");
				//Console.WriteLine($"current time : {DateTime.Now}");

				if (!isTokenExpired)
				{
					// Check buffer time (less than 1 hour?)
					var bufferTime = (authTokenData.AuthTokenCreatedDatetime?.AddSeconds(Convert.ToDouble(authTokenData.Expiry)) - DateTime.Now );

                    //var bufferTime = (authTokenData.AuthTokenCreatedDatetime?.AddMinutes(Convert.ToDouble(authTokenData.Expiry)) - DateTime.Now);
                    //Console.WriteLine($"Buffer Time: {bufferTime}");
					//Console.ReadKey();

					if (bufferTime <= TimeSpan.FromHours(1))
					{

						#region 3rd API for refresh token

						// Parametres
						parameters = "";

						// Headers
						httpClient = _httpClientFactory.CreateClient();
                        httpClient.DefaultRequestHeaders.Add("clientid", clientid);
                        httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
                        httpClient.DefaultRequestHeaders.Add("state-cd", statecd);
                        httpClient.DefaultRequestHeaders.Add("ip-usr", ipusr);
                        httpClient.DefaultRequestHeaders.Add("txn", txn);

                        // Body
                        var RequestBody3 = new
                        {
                            action = "REFRESHTOKEN",
                            username = userName,
                            auth_token = authTokenData.AuthToken,
							key = authTokenData.XAppKey,
							sek = authTokenData.SEK
                        };
                        var content3 = new StringContent(JsonConvert.SerializeObject(RequestBody3), Encoding.UTF8, "application/json");

                        // Url
                        baseUrl = _configuration["PurchaseRegisterIRISAPI:BaseUrl"];
                        AuthUrl = _configuration["PurchaseRegisterIRISAPI:RefreshAuthToken"];
						apiUrl = $"{baseUrl}{AuthUrl}";

                        // API call - Response
                        var response3 = await httpClient.PostAsync(apiUrl, content3);
                        var responseData3 = await response3.Content.ReadAsStringAsync();
                        var result3 = JsonConvert.DeserializeObject<JObject>(responseData3);

						//Console.WriteLine($"Response 3 : {response3}");
						//Console.WriteLine($"Result 3 : {result3}");
						//Console.ReadKey();

						// save API data to database
						var headersDict3 = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                        headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict3);
                        APIsDataModel ApiData3 = new APIsDataModel();
                        {
                            ApiData3.ClientGstin = ClientGSTIN;
                            ApiData3.RequestNumber = ticketnumber;
                            ApiData3.SessionID = sessionId;
                            ApiData3.RequestURL = apiUrl;
                            ApiData3.RequestParameters = parameters;
                            ApiData3.RequestHeaders = headersJson;
                            ApiData3.RequestBody = JsonConvert.SerializeObject(RequestBody3);
                            ApiData3.Response = responseData3;
                            ApiData3.ResponseCode = $"{(int)response3.StatusCode} {response3.StatusCode}";
                            ApiData3.Status = getStatus((int)response3.StatusCode);
                        };
                        await _gSTR2DataBusiness.savePortalAPIsData(ApiData3);

                        // validate response
                        if (result3["status_cd"].ToString() != "1")
						{
							throw new Exception("Refresh Auth Token API Call - Failed to retrieve 'auth_token' from the response.");
						}

                        string authToken = result3["auth_token"].ToString();
                        string expiry = result3["expiry"].ToString();
                        string sek = result3["sek"].ToString();

                        // Save token info  into DB
                        await _gSTR2DataBusiness.updateTokenData(new GSTR2TokenDataModel
                        {
                            AuthToken = authToken,
                            Expiry = expiry,
                            SEK = sek
                        });

                        #endregion

                    }

                    // Token is still valid, proceed to 4th API (later step)
					//Console.WriteLine("Token is still valid. Proceeding to 4th API call.");
                    return Json(new { success = true, message = "Token is valid. Continue." });
					
				}

                #region Request for OTP         1st Api call 
                // Parameters
                parameters = "";

                // Headers
                httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Add("clientid", clientid);
                httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
                httpClient.DefaultRequestHeaders.Add("state-cd", statecd);
                httpClient.DefaultRequestHeaders.Add("ip-usr", ipusr);
                httpClient.DefaultRequestHeaders.Add("txn", txn);

                // Body
                var RequestBody1 = new
                {
                    action = "OTPREQUEST",
                    username = userName
                };
                var content = new StringContent(JsonConvert.SerializeObject(RequestBody1), Encoding.UTF8, "application/json");

                // Url
                baseUrl = _configuration["PurchaseRegisterIRISAPI:BaseUrl"];
                AuthUrl = _configuration["PurchaseRegisterIRISAPI:RequestForOTP"];
                apiUrl = $"{baseUrl}{AuthUrl}";

                // API call - Response
                var response = await httpClient.PostAsync(apiUrl, content);
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);
                //Console.WriteLine($"Response 1 : {response}");
                //Console.WriteLine($"Result 1 : {result}");
                //Console.ReadKey();

                // save API data to database   APIsDataModel
                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketnumber;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody1);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                };
                await _gSTR2DataBusiness.savePortalAPIsData(ApiData);
                //Console.WriteLine($"Client Gstin : {ClientGSTIN} ");
                //Console.WriteLine($"Request Number : {ticketnumber}");
                //Console.WriteLine($"Session ID : {sessionId}");
                //Console.WriteLine($"Request URL : {apiUrl}");
                //Console.WriteLine($"Request Parameters : {parameters}");
                //Console.WriteLine($"Request Headers: {headersJson}");
                //Console.WriteLine($"Request Body : {JsonConvert.SerializeObject(RequestBody1)}");
                //Console.WriteLine($"Response : {responseData}");
                //Console.WriteLine($"Response Code : {(int)response.StatusCode} {response.StatusCode}");
                //Console.WriteLine($"Status : {getStatus((int)response.StatusCode)}");

                //return Json(new
                //{
                //    success = false,
                //    askForOtp = true
                //});

                if (result["status_cd"].ToString() != "1")
                {
                    return Json(new
                    {
                        success = false,
                        askForOtp = false,
                        message = "OTP Request API Call - Failed to retrieve 'X-App-Key' from the response."
                    });
                    //throw new Exception("1st API Call - Failed to retrieve 'X-App-Key' from the response.");
                }

                //if (!xAppKeyResponse.IsSuccess)
                //    return Json(new { success = false, message = "Failed to get x-app-key from 1st API" });

                string x_app_key = result["x-app-key"].ToString();
                //Console.WriteLine($"x_app_key : {x_app_key}");

                #endregion

                // Return to front-end to prompt for OTP
                return Json(new
				{
					success = false,
					askForOtp = true,
					xAppKey = x_app_key
                });
            }
            catch (Exception ex)
            {
                string ErrorMessage = $"Error: {ex.Message}";
                return Json(new { message = ErrorMessage });
            }

            #endregion
            
		}

        [HttpPost]
        public async Task<IActionResult> SubmitOtpAndContinue(string ClientGSTIN, string ticketNo, string otp, string xAppKey)
        {
			//Console.WriteLine("HI");
            string sessionId = HttpContext.Session.Id;
            ViewBag.Messages = "Admin";

            var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);

            try
            {
                #region Request For AuthToken            2nd api call
                string userName = UserAPIData.GstPortalUsername;
                string clientid = _configuration["PurchaseRegisterIRISAPI:ClientId"];
                string clientSecret = _configuration["PurchaseRegisterIRISAPI:ClientSecret"];
                string statecd = ClientGSTIN.Substring(0, 2);  // Extracting state code from GSTIN
                string ipusr = _configuration["PurchaseRegisterIRISAPI:ipurs"];
                string txn = _configuration["PurchaseRegisterIRISAPI:txn"];

                // Parameters
                string parameters = "";

                // Headers
                var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Add("clientid", clientid);
                httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
                httpClient.DefaultRequestHeaders.Add("state-cd", statecd);
                httpClient.DefaultRequestHeaders.Add("ip-usr", ipusr);
                httpClient.DefaultRequestHeaders.Add("txn", txn);

                // Body
                var RequestBody1 = new
                {
                    action = "AUTHTOKEN",
                    username = userName,
                    otp = otp,
					key = xAppKey
                };
                var content = new StringContent(JsonConvert.SerializeObject(RequestBody1), Encoding.UTF8, "application/json");

                // Url
                string baseUrl = _configuration["PurchaseRegisterIRISAPI:BaseUrl"];
                string AuthUrl = _configuration["PurchaseRegisterIRISAPI:RequestForAuthToken"];
                string apiUrl = $"{baseUrl}{AuthUrl}";

                // API call - Response
                var response = await httpClient.PostAsync(apiUrl, content);
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);
                //Console.WriteLine($"Response 2 : {response}");
                //Console.WriteLine($"Result 2 : {result}");
                //Console.ReadKey();

                // save API data to database   APIsDataModel
                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                string headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketNo;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody1);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                };
                await _gSTR2DataBusiness.savePortalAPIsData(ApiData);

				//return Json(new
				//{
				//	success = false,
				//	askAgain = true,
				//	message = "Invalid OTP. Please enter again."
				//});

				// validate response
				if (result["status_cd"].ToString() != "1")
				{
                    if (result["error"]["error_cd"].ToString() == "AUTH4033")
                    {
                        return Json(new
                        {
                            success = false,
                            askAgain = true,
                            message = "Invalid OTP. Please enter again."
                        });
                    }
                    return Json(new { success = false, message = "Auth Token API Call - Failed to retrieve 'Token' from the response." });

                }

                string authToken = result["auth_token"].ToString();
				string expiry = result["expiry"].ToString();
				string sek = result["sek"].ToString();

                // Step: Save token info
                await _gSTR2DataBusiness.saveTokenData(new GSTR2TokenDataModel
                {
                    ClientGstin = ClientGSTIN,
                    RequestNumber = ticketNo,
                    UserName = userName,
                    XAppKey = xAppKey,
                    OTP = otp,
                    AuthToken = authToken,
                    Expiry = expiry,
                    SEK = sek
                });
                #endregion

                return Json(new { success = true, message = "OTP Verified. Token saved." });

            }
            catch (Exception ex)
            {
                string ErrorMessage = $"Error: {ex.Message}";
                return Json(new { message = ErrorMessage });
            }

        }

		public async Task<IActionResult> ContinueWith4thApi(string ticketnumber, string clientGSTIN)
		{
            //Console.WriteLine($"TicketNumber  4: {ticketnumber}");
            //Console.WriteLine($"ClientGSTIN 4: {clientGSTIN}");

            string sessionId = HttpContext.Session.Id;
			ViewBag.Messages = "Admin";
			var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, clientGSTIN);
			ViewBag.Ticket = Ticket;

			var authTokenData = await _gSTR2DataBusiness.GetTokenDataAsync(clientGSTIN);

			DataTable PRDataTable = new DataTable();
			// Read headers
			var EInvoiceColumns = _configuration["Portal"].Split(',').Select(x => x.Trim()).ToList();
			foreach (string header in EInvoiceColumns)
			{
				PRDataTable.Columns.Add(header.Trim(), typeof(string));
			}



			string txnPeriod = Ticket.TxnPeriod;
			string formattedPeriod = DateTime.ParseExact(txnPeriod, "MMM-yy", CultureInfo.InvariantCulture)
								  .ToString("MMyyyy");


			try
			{
				#region 4th api call 
				string clientid = _configuration["PurchaseRegisterIRISAPI:ClientId"];
				string clientSecret = _configuration["PurchaseRegisterIRISAPI:ClientSecret"];
				string statecd = clientGSTIN.Substring(0, 2);  // Extracting state code from GSTIN
				string ipusr = _configuration["PurchaseRegisterIRISAPI:ipurs"];
				string txn = _configuration["PurchaseRegisterIRISAPI:txn"];

				// Parameters
				string action = "GETINV";
				string gstin = clientGSTIN;
				string section = "B2B";

				string parameter = $"action={action}&gstin={gstin}&section={section}";

				// Headers
				var httpClient = _httpClientFactory.CreateClient();
				httpClient.DefaultRequestHeaders.Add("clientid", clientid);
				httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
				httpClient.DefaultRequestHeaders.Add("state-cd", statecd);
				httpClient.DefaultRequestHeaders.Add("ip-usr", ipusr);
				httpClient.DefaultRequestHeaders.Add("txn", txn);
				httpClient.DefaultRequestHeaders.Add("auth-token", authTokenData.AuthToken);
				httpClient.DefaultRequestHeaders.Add("username", authTokenData.UserName);
				httpClient.DefaultRequestHeaders.Add("gstin", authTokenData.ClientGstin);
				httpClient.DefaultRequestHeaders.Add("ret-period", formattedPeriod);
				httpClient.DefaultRequestHeaders.Add("x-sek", authTokenData.SEK);
				httpClient.DefaultRequestHeaders.Add("x-app-key", authTokenData.XAppKey);

				// Body
				var RequestBody1 = new
				{

				};
				var content = new StringContent(JsonConvert.SerializeObject(RequestBody1), Encoding.UTF8, "application/json");

				// Url
				string baseUrl = _configuration["PurchaseRegisterIRISAPI:BaseUrl"];
				string AuthUrl = _configuration["PurchaseRegisterIRISAPI:GetPortalData"];
				string apiUrl = $"{baseUrl}{AuthUrl}?{parameter}";

				// Api call - Response
				var response = await httpClient.GetAsync(apiUrl);
				var responseData = await response.Content.ReadAsStringAsync();
				var result = JsonConvert.DeserializeObject<JObject>(responseData);

				// save API data to database
				var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
				string headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
				APIsDataModel ApiData = new APIsDataModel
				{
					ClientGstin = clientGSTIN,
					RequestNumber = ticketnumber,
					SessionID = sessionId,
					RequestURL = apiUrl,
					RequestParameters = parameter,
					RequestHeaders = headersJson,
					RequestBody = JsonConvert.SerializeObject(RequestBody1),
					Response = responseData,
					ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}",
					Status = getStatus((int)response.StatusCode)
				};
				await _gSTR2DataBusiness.savePortalAPIsData(ApiData);

				//Console.WriteLine($"Response 4 : {response}");
				//Console.WriteLine($"Result 4 : {result}");

				// validate response

				// if failed , return error message
				if (result["b2b"] == null)
				{
                    //{
                    //	"est": "30",
                    //	"token": "2ffea879af264c36915fc7d5ba381a36"
                    //}
                    if (result["token"] != null)
					{
                        string token = result["token"]?.ToString();

                        return Json(new { success = true, message = "Try after sometime", token = token });
                        //return RedirectToAction("ContinueWith5thApi", new { ticketnumber, clientGSTIN, token5 });

                    }
                    //{
                    //	"status_code": 0,
                    //	"message": "The file for the request date and type is already generated on 13/08/2025 11:37:17. To download the same file, use token 6001f2a9df354b2ca15532fc86c1d4ae. The link is valid till 1 day."
                    //}
                    else if (result["message"] != null)
                    {
                        string message = result["message"]?.ToString();

                        // Match any non-space sequence after the word "token"
                        var match = System.Text.RegularExpressions.Regex.Match(message, @"token\s+([^\s.]+)", RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            string tokenFromMessage = match.Groups[1].Value;
                            return Json(new { success = true, token = tokenFromMessage });
                        }
                    }
                    var apiMessage = result["message"]?.ToString();
                    string msg = $"\"GET INVOICES\" API Call Failed - Response - \"{apiMessage}\"";

                    return Json(new { failure = true, message = msg });

                }

                // convert response data to DataTable
                PRDataTable = ConvertJsontoPRDataTable(responseData, clientGSTIN, PRDataTable);

				#endregion

				// Save the data to database
				await _gSTR2DataBusiness.SaveGSTR2DataAsync(PRDataTable, ticketnumber, clientGSTIN);

			}
			catch (Exception ex)
			{
				string Message = $"Error: {ex.Message}";
                return Json(new { failure = true, message = Message });
                //return View("/Views/Admin/CompareGstFiles/CompareGSTAPI_IRIS.cshtml");
			}

			return Json( new { goToCompare = true });
		}

        public async Task<IActionResult> ContinueWith5thApi(string ticketnumber, string clientGSTIN ,string token5)
		{
            string sessionId = HttpContext.Session.Id;
            ViewBag.Messages = "Admin";
            var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, clientGSTIN);
            ViewBag.Ticket = Ticket;

            var authTokenData = await _gSTR2DataBusiness.GetTokenDataAsync(clientGSTIN);

            DataTable PRDataTable = new DataTable();
            // Read headers
            var EInvoiceColumns = _configuration["Portal"].Split(',').Select(x => x.Trim()).ToList();
            foreach (string header in EInvoiceColumns)
            {
                PRDataTable.Columns.Add(header.Trim(), typeof(string));
            }


            string txnPeriod = Ticket.TxnPeriod;
            string formattedPeriod = DateTime.ParseExact(txnPeriod, "MMM-yy", CultureInfo.InvariantCulture)
                                  .ToString("MMyyyy");
			//Console.WriteLine($"TicketNumber 5: {ticketnumber}");
			//Console.WriteLine($"ClientGSTIN 5: {clientGSTIN}");



			try
			{

				#region  5th API call to get data
				string clientid = _configuration["PurchaseRegisterIRISAPI:ClientId"];
				string clientSecret = _configuration["PurchaseRegisterIRISAPI:ClientSecret"];
				string statecd = clientGSTIN.Substring(0, 2);  // Extracting state code from GSTIN
				string ipusr = _configuration["PurchaseRegisterIRISAPI:ipurs"];
				string txn = _configuration["PurchaseRegisterIRISAPI:txn"];

				// Parameters
				string action5 = "FILEDET";
				string gstin5 = clientGSTIN;
				string parameter5 = $"action={action5}&gstin={gstin5}&token={token5}";

				// Headers
				var httpClient = _httpClientFactory.CreateClient();
				httpClient.DefaultRequestHeaders.Add("clientid", clientid);
				httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
				httpClient.DefaultRequestHeaders.Add("state-cd", statecd);
				httpClient.DefaultRequestHeaders.Add("ip-usr", ipusr);
				httpClient.DefaultRequestHeaders.Add("txn", txn);
				httpClient.DefaultRequestHeaders.Add("auth-token", authTokenData.AuthToken);
				httpClient.DefaultRequestHeaders.Add("username", authTokenData.UserName);
				httpClient.DefaultRequestHeaders.Add("gstin", authTokenData.ClientGstin);
				httpClient.DefaultRequestHeaders.Add("ret-period", formattedPeriod);
				httpClient.DefaultRequestHeaders.Add("x-sek", authTokenData.SEK);
				httpClient.DefaultRequestHeaders.Add("x-app-key", authTokenData.XAppKey);

				// Body
				var RequestBody5 = new
				{

				};
				var content5 = new StringContent(JsonConvert.SerializeObject(RequestBody5), Encoding.UTF8, "application/json");

				// Url
				string baseUrl5 = _configuration["PurchaseRegisterIRISAPI:BaseUrl"];
				string AuthUrl5 = _configuration["PurchaseRegisterIRISAPI:GetPortalData"];
				string apiUrl5 = $"{baseUrl5}{AuthUrl5}?{parameter5}";

				// Api call - Response
				var response5 = await httpClient.GetAsync(apiUrl5);
				var responseData5 = await response5.Content.ReadAsStringAsync();
				var result5 = JsonConvert.DeserializeObject<JObject>(responseData5);

				// save API data to database
				var headersDict5 = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
				string headersJson5 = System.Text.Json.JsonSerializer.Serialize(headersDict5);
				APIsDataModel ApiData5 = new APIsDataModel
				{
					ClientGstin = clientGSTIN,
					RequestNumber = ticketnumber,
					SessionID = sessionId,
					RequestURL = apiUrl5,
					RequestParameters = parameter5,
					RequestHeaders = headersJson5,
					RequestBody = JsonConvert.SerializeObject(RequestBody5),
					Response = responseData5,
					ResponseCode = $"{(int)response5.StatusCode} {response5.StatusCode}",
					Status = getStatus((int)response5.StatusCode)
				};
				await _gSTR2DataBusiness.savePortalAPIsData(ApiData5);

				//Console.WriteLine($"Response 5 : {response5}");
				Console.WriteLine($"Result 5 : {result5}");
				Console.WriteLine($"1 - {result5["status_code"]?.ToString() == "0"} , 2 - {result5["urls"] != null}");


				//Result 5 : {
				//           "status_code": 0,
				//           "message": "File Generation is in progress, please try after sometime. "
				//           }
				if (result5["status_code"]?.ToString() == "0")
				{
					string message = result5["message"]?.ToString();

					if (!string.IsNullOrEmpty(message) && message.Contains("File Generation"))
					{

						return Json(new { message = "Try after sometime", token = token5 });
						//return View("/Views/Admin/CompareGstFiles/CompareGSTAPI_IRIS.cshtml");
					}
					else
					{
                        return Json(new
                        {
                            error = true,
                            message = $"\"Get IMS File Details\" API Call Failed - Response - \"{message}\"",
                        });
                        //throw new Exception($"5th API Call - Failed due to {message}");
					}
				}

				//Result 5 : {
				//            "urls": [
				//                  {
				//                    "ul": "https://uatfiles.gst.gov.in/imsreturns/imsJsn/27CKGPR5841G1ZU_B2B_2025-06-30_IMS_1.tar.gz?md5=Gi7S2f5RDtkxQSSFexytiw&expires=1751366867",
				//                    "ic": 100,
				//                    "hash": "d6caaf367f132aa4461450fb0f63139e758c9d9f760eaadd99cf1f686ece5337"
				//                  },
				//                  {
				//                    "ul": "https://uatfiles.gst.gov.in/imsreturns/imsJsn/27CKGPR5841G1ZU_B2B_2025-06-30_IMS_2.tar.gz?md5=oI3LkTjnUgofliAM85zAfg&expires=1751366867",
				//                    "ic": 15,
				//                    "hash": "76d1431b724ec4406d246eb714a5db276107c1f221dfbdbfec13ffee3f5a42b0"
				//                  }
				//            ],
				//            "ek": "ybANSSYD4c1AUxjfleN9X3x/bl8Z0Hq8kK+KJNf7rzw=",
				//            "fc": 2
				//            }
				else if (result5["urls"] != null)
				{
					string ek = result5["ek"]?.ToString();
					int fileCount = (int)result5["fc"];
					JArray URLs = (JArray)result5["urls"];
					string[] urls = URLs.Select(u => u["ul"]?.ToString()).ToArray();
					// replace uat to sb in each url
					urls = urls.Select(u => u.Replace("uatfiles.gst.gov.in", "sbfiles.gst.gov.in")).ToArray();
					foreach (string url in urls)
					{
						Console.WriteLine($"URL: {url}");
						// Step 1: Download tar.gz file
						using HttpClient client = new HttpClient();
						byte[] fileBytes = await client.GetByteArrayAsync(url);

						// Step 2: Save it temporarily
						string tempTarGzPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".tar.gz");
						await System.IO.File.WriteAllBytesAsync(tempTarGzPath, fileBytes);

						// Step 3: Extract .tar.gz
						string extractedFolderPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
						Directory.CreateDirectory(extractedFolderPath);

						// First extract .gz to .tar
						string tempTarPath = Path.ChangeExtension(tempTarGzPath, ".tar");
						using (FileStream originalFileStream = new FileStream(tempTarGzPath, FileMode.Open, FileAccess.Read))
						using (FileStream decompressedFileStream = new FileStream(tempTarPath, FileMode.Create))
						using (GZipStream decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress))
						{
							decompressionStream.CopyTo(decompressedFileStream);
						}

						// Now extract .tar file
						using (var archive = TarArchive.Open(tempTarPath))
						{
							foreach (var entry in archive.Entries.Where(entry => !entry.IsDirectory))
							{
								string extractedFilePath = Path.Combine(extractedFolderPath, entry.Key);
								entry.WriteToFile(extractedFilePath);

								// Step 4: Read JSON file content
								string jsonContent = await System.IO.File.ReadAllTextAsync(extractedFilePath);

								// Step 5: Pass to DecryptGSTNResponse
								PRDataTable = DecryptGSTNResponse(jsonContent, ek, clientGSTIN, PRDataTable);
							}
						}

						// Optional: Clean up temp files
						System.IO.File.Delete(tempTarGzPath);
						System.IO.File.Delete(tempTarPath);
						Directory.Delete(extractedFolderPath, true);

						//Console.WriteLine("Hi");
					}

					// Save the data to database
					await _gSTR2DataBusiness.SaveGSTR2DataAsync(PRDataTable, ticketnumber, clientGSTIN);

					return Json(new { success = true, ticketno = ticketnumber, gstin = clientGSTIN });
				}

				// Result 5 : {
				//    "error_code": 500,
				//	"message": "Internal Server Error",
				//	"description": "An unexpected condition was encountered. Our service team has been dispatched to bring it back online."
				//}

				else if (result5["error_code"]?.ToString() == "500")
				{
					string message = result5["message"]?.ToString();
					message += result5["description"]?.ToString();
					return Json(new
					{
						error = true,
						message = message,
					});
				}
                #endregion

            }
            catch (Exception ex)
			{
                return Json(new
                {
                    error = true,
                    message = $"5th API Call - Failed due to {ex}",
                });
                //ViewBag.ErrorMessage = $"Error: {ex.Message}";
                //return View("/Views/Admin/CompareGstFiles/CompareGSTAPI_IRIS.cshtml");
            }

            return RedirectToAction("CompareDataTables", new { ticketnumber, clientGSTIN });
        }
       
		public async Task<IActionResult> CompareDataTables(string ticketnumber, string clientGSTIN)
        {
            ViewBag.Messages = "Admin";
            var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, clientGSTIN);
            ViewBag.Ticket = Ticket;

            //Console.WriteLine($"TicketNumber Compare: {ticketnumber}");
			//Console.WriteLine($"ClientGSTIN Compare: {clientGSTIN}");
            // Comparison logic (unchanged)
            DataTable gstr2Data = await _gSTR2DataBusiness.GetPortalDataBasedOnTicketAsync(ticketnumber, clientGSTIN);
			//_logger.LogInformation("GSTR2 Columns: " + string.Join(", ", gstr2Data.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
			//_logger.LogInformation("Total GSTR2 Data Rows: " + gstr2Data.Rows.Count);

			DataTable purchaseData = await _purchaseDataBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, clientGSTIN);
			//_logger.LogInformation("purchaseData " + string.Join(", ", purchaseData.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
			//_logger.LogInformation("Total Purchase Data Rows: " + purchaseData.Rows.Count);

			//Console.ReadKey();
			await _compareGstBusiness.SaveCompareDataAsync(purchaseData, gstr2Data);

            string AdminFileName = $"{ticketnumber}_API_IRIS.csv";
            await _purchaseTicketBusiness.UpdatePurchaseTicketAsync(ticketnumber, clientGSTIN, AdminFileName);

            var data = await _compareGstBusiness.GetComparedDataBasedOnTicketAsync(ticketnumber, clientGSTIN);
            ViewBag.ReportDataList = data;
            //_logger.LogInformation("Total Compared Data Rows: " + data.Count);

            // store number and sum in a model and save it in 
            ViewBag.Summary = GenerateSummary(data);
            var summary = ViewBag.Summary;
            //_logger.LogInformation($"invoice 1 count {summary.catagory1InvoiceNumber},portal 1 count {summary.catagory1PortalNumber}");

            return View("/Views/Admin/CompareGstFiles/Compare.cshtml");
        }

        private DataTable DecryptGSTNResponse(string response, string ek,string userGstin, DataTable PRDataTable)
        {
            // Decode the key
            byte[] key = Convert.FromBase64String(ek);

                // Test base64 decoding
                Convert.FromBase64String(response); 
                
                // Decrypt the response
                byte[] resul = AesDecryptWithKey(response, key);
                //Console.WriteLine("1");

                // Decode the result to a UTF-8 string
                string jsonStr = Encoding.UTF8.GetString(resul);
                //Console.WriteLine("2");


                // Additional base64 decoding of the decrypted string
                jsonStr = Encoding.UTF8.GetString(Convert.FromBase64String(jsonStr));

                // Try parsing JSON
                JsonDocument jsonData = JsonDocument.Parse(jsonStr);

				string jsonPreview = System.Text.Json.JsonSerializer.Serialize(jsonData, new JsonSerializerOptions { WriteIndented = true });


                var result = JsonConvert.DeserializeObject<JObject>(jsonPreview);
                // Check if 'b2b' exists
                if (result["b2b"] != null && result["b2b"].Type == JTokenType.Array)
                {
                    var b2bArray = (JArray)result["b2b"];

                    if (b2bArray.Count > 0)
                    {
                        foreach (var item in b2bArray)
                        {
                            DataRow row = PRDataTable.NewRow();

                            row["User GSTIN"] = userGstin; // or your value
                            row["GstRegType"] = "";
                            row["invoiceno"] = item["inum"]?.ToString();
                            row["invoice_date"] = item["idt"]?.ToString();
                            row["Supplier GSTIN"] = item["stin"]?.ToString();
                            row["SupplierName"] = "";
                            row["IsRcmApplied"] = "N";
                            row["InvoiceValue"] = item["val"]?.ToString();
                            row["ItemTaxableValue"] = item["txval"]?.ToString();
                            row["GstRate"] = "";
                            row["IGSTAmount"] = item["iamt"]?.ToString();
                            row["CGSTAmount"] = item["camt"]?.ToString();
                            row["SGSTAmount"] = item["samt"]?.ToString();
                            row["CESS"] = item["cess"]?.ToString();
                            row["IsReturnFiled"] = item["srcfilstatus"]?.ToString();
                            row["ReturnPeriod"] = item["rtnprd"]?.ToString();


                            //{     // Sample API Response
                            //    "val": 8136.1,
                            //    "samt": 620.55,
                            //    "txval": 6895,
                            //    "rtnprd": "052025",
                            //    "camt": 620.55,
                            //    "inum": "009",
                            //    "cess": 0,
                            //    "srcfilstatus": "Filed",
                            //    "iamt": 0,
                            //    "stin": "33IDCPS9065J1ZY",
                            //    "srcform": "R1",
                            //    "inv_typ": "R",
                            //    "pos": "33",
                            //    "idt": "21-05-2025",
                            //    "action": "N",
                            //    "ispendactblocked": "N",
                            //    "hash": "de46f7e9d414b36ab6fb41552cfbcd5dc1ba7680ecb78626bb39634db5f89e48"
                            //},


                            PRDataTable.Rows.Add(row);
                        }
                    }
                }
           
            return PRDataTable;
        }
       
		static byte[] AesDecryptWithKey(string message, byte[] key)
        {
            try
            {
                // Decode the base64-encoded message
                byte[] messageBytes = Convert.FromBase64String(message);

                // Initialize AES cipher in ECB mode
                using (Aes aes = Aes.Create())
                {
                    aes.Key = key;
                    aes.Mode = CipherMode.ECB;
                    aes.Padding = PaddingMode.PKCS7; // Equivalent to Python's PKCS7 padding

                    // Create decryptor
                    using (ICryptoTransform decryptor = aes.CreateDecryptor())
                    {
                        // Decrypt the message
                        return decryptor.TransformFinalBlock(messageBytes, 0, messageBytes.Length);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Decryption failed: {ex.Message}");
            }
        }

        private DataTable ConvertJsontoPRDataTable(string responseData, string userGstin, DataTable PRDataTable)
		{
            var result = JsonConvert.DeserializeObject<JObject>(responseData);
            // Check if 'b2b' exists
            if (result["b2b"] != null && result["b2b"].Type == JTokenType.Array)
            {
                var b2bArray = (JArray)result["b2b"];

                if (b2bArray.Count > 0)
                {
                    foreach (var item in b2bArray)
                    {
                        DataRow row = PRDataTable.NewRow();

                        row["User GSTIN"] = userGstin; // or your value
                        row["GstRegType"] = "";
                        row["invoiceno"] = item["inum"]?.ToString();
                        row["invoice_date"] = item["idt"]?.ToString();
                        row["Supplier GSTIN"] = item["stin"]?.ToString();
                        row["SupplierName"] = "";
                        row["IsRcmApplied"] = "N";
                        row["InvoiceValue"] = item["val"]?.ToString();
                        row["ItemTaxableValue"] = item["txval"]?.ToString();
                        row["GstRate"] = "";
                        row["IGSTAmount"] = item["iamt"]?.ToString();
                        row["CGSTAmount"] = item["camt"]?.ToString();
                        row["SGSTAmount"] = item["samt"]?.ToString();
                        row["CESS"] = item["cess"]?.ToString();
                        row["IsReturnFiled"] = item["srcfilstatus"]?.ToString();
                        row["ReturnPeriod"] = item["rtnprd"]?.ToString();


						//{     // Sample API Response
						//    "val": 8136.1,
						//    "samt": 620.55,
						//    "txval": 6895,
						//    "rtnprd": "052025",
						//    "camt": 620.55,
						//    "inum": "009",
						//    "cess": 0,
						//    "srcfilstatus": "Filed",
						//    "iamt": 0,
						//    "stin": "33IDCPS9065J1ZY",
						//    "srcform": "R1",
						//    "inv_typ": "R",
						//    "pos": "33",
						//    "idt": "21-05-2025",
						//    "action": "N",
						//    "ispendactblocked": "N",
						//    "hash": "de46f7e9d414b36ab6fb41552cfbcd5dc1ba7680ecb78626bb39634db5f89e48"
						//},


                        PRDataTable.Rows.Add(row);
                    }
                }
            }

			return PRDataTable;
        }

     
        #endregion

        #region Helper methods to convert Data(Json to CSV, CSV To DataTable
        private DataTable ConvertIRISJsonToDataTable(string jsonData, string Gstin)
		{
			// Deserialize the root object (assumed to be an object containing multiple arrays like "b2b", "b2ba", etc.)
			var rootObject = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonData);

			// Create a DataTable
			DataTable dt = new DataTable();

			// Define the headers as per your requirement
			dt.Columns.Add("User GSTIN");
			dt.Columns.Add("GstRegType");
			dt.Columns.Add("invoiceno");
			dt.Columns.Add("invoice_date");
			dt.Columns.Add("Supplier GSTIN");
			dt.Columns.Add("SupplierName");
			dt.Columns.Add("IsRcmApplied");
			dt.Columns.Add("InvoiceValue");
			dt.Columns.Add("ItemTaxableValue");
			dt.Columns.Add("GstRate");
			dt.Columns.Add("IGSTAmount");
			dt.Columns.Add("CGSTAmount");
			dt.Columns.Add("SGSTAmount");
			dt.Columns.Add("CESS");
			dt.Columns.Add("IsReturnFiled");
			dt.Columns.Add("ReturnPeriod");

			// Loop through each key in the root object
			foreach (var key in rootObject.Keys)
			{
				// Deserialize the array for each key (e.g., "b2b", "b2ba", etc.)
				var dataArray = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(rootObject[key].ToString());

				// Loop through the dataArray (which contains the data for each key)
				foreach (var item in dataArray)
				{
					DataRow row = dt.NewRow();

					// Add data from the item dictionary into the row
					row["User GSTIN"] = Gstin;
					row["GstRegType"] = key; // item.ContainsKey("key") ? item["key"] : DBNull.Value;
					row["invoiceno"] = item.ContainsKey("inum") ? item["inum"] : (item.ContainsKey("nt_num") ? item["nt_num"] : DBNull.Value);
					row["invoice_date"] = item.ContainsKey("idt") ? item["idt"] : (item.ContainsKey("nt_dt") ? item["nt_dt"] : DBNull.Value);
					row["Supplier GSTIN"] = item.ContainsKey("stin") ? item["stin"] : DBNull.Value;
					row["SupplierName"] = DBNull.Value; // Assuming no field for SupplierName in the JSON
					row["IsRcmApplied"] = DBNull.Value; // Assuming no field for IsRcmApplied in the JSON
					row["InvoiceValue"] = item.ContainsKey("val") ? item["val"] : DBNull.Value;
					row["ItemTaxableValue"] = item.ContainsKey("txval") ? item["txval"] : DBNull.Value;

					// If "GstRate" is not available, calculate it
					decimal gstRate = 0;
					if (item.ContainsKey("iamt") && item.ContainsKey("txval"))
					{
						decimal igst = Convert.ToDecimal(item["iamt"]);
						decimal taxableValue = Convert.ToDecimal(item["txval"]);
						gstRate = (igst / taxableValue) * 100;
					}
					row["GstRate"] = DBNull.Value;  //gstRate;

					row["IGSTAmount"] = item.ContainsKey("iamt") ? item["iamt"] : DBNull.Value;
					row["CGSTAmount"] = item.ContainsKey("camt") ? item["camt"] : DBNull.Value;
					row["SGSTAmount"] = item.ContainsKey("samt") ? item["samt"] : DBNull.Value;
					row["CESS"] = item.ContainsKey("cess") ? item["cess"] : DBNull.Value;
					row["IsReturnFiled"] = item.ContainsKey("srcfilstatus") ? item["srcfilstatus"] : DBNull.Value;
					row["ReturnPeriod"] = item.ContainsKey("rtnprd") ? item["rtnprd"] : DBNull.Value;

					// Add the row to DataTable
					dt.Rows.Add(row);
				}
			}

			return dt;
		}
		private DataTable ConvertJsonToDataTable(string jsonData, string userGstin)
		{
			DataTable dt = new DataTable();

			// Define columns
			dt.Columns.Add("User GSTIN");
			dt.Columns.Add("GstRegType");
			dt.Columns.Add("invoiceno");
			dt.Columns.Add("invoice_date");
			dt.Columns.Add("SupplierGSTIN");
			dt.Columns.Add("SupplierName");
			dt.Columns.Add("IsRcmApplied");
			dt.Columns.Add("InvoiceValue", typeof(decimal));
			dt.Columns.Add("ItemTaxableValue", typeof(decimal));
			dt.Columns.Add("GstRate", typeof(decimal));
			dt.Columns.Add("IGSTAmount", typeof(decimal));
			dt.Columns.Add("CGSTAmount", typeof(decimal));
			dt.Columns.Add("SGSTAmount", typeof(decimal));
			dt.Columns.Add("CESS", typeof(decimal));
			dt.Columns.Add("IsReturnFiled");
			dt.Columns.Add("ReturnPeriod");

			JObject obj = JObject.Parse(jsonData);
			JArray b2b = (JArray)obj["data"]?["b2b"];

			foreach (var party in b2b)
			{
				string ctin = party.Value<string>("ctin");
				string ret_period = party.Value<string>("flprdr1");

				foreach (var inv in party["inv"])
				{
					string inum = inv.Value<string>("inum");
					string idt = inv.Value<string>("idt");
					string inv_typ = inv.Value<string>("inv_typ");
					string rchrg = inv.Value<string>("rchrg");
					decimal val = inv.Value<decimal?>("val") ?? 0;

					foreach (var item in inv["itms"])
					{
						var itm_det = item["itm_det"];
						decimal txval = itm_det.Value<decimal?>("txval") ?? 0;
						decimal rt = itm_det.Value<decimal?>("rt") ?? 0;
						decimal iamt = itm_det.Value<decimal?>("iamt") ?? 0;
						decimal camt = itm_det.Value<decimal?>("camt") ?? 0;
						decimal samt = itm_det.Value<decimal?>("samt") ?? 0;
						decimal csamt = itm_det.Value<decimal?>("csamt") ?? 0;

						dt.Rows.Add(
							userGstin,         // User GSTIN
							inv_typ,           // GstRegType
							inum,              // invoiceno
							idt,               // invoice_date
							ctin,              // SupplierGSTIN
							"",                // SupplierName (Not provided)
							rchrg,             // IsRcmApplied
							val,               // InvoiceValue
							txval,             // ItemTaxableValue
							rt,                // GstRate
							iamt,              // IGSTAmount
							camt,              // CGSTAmount
							samt,              // SGSTAmount
							csamt,             // CESS
							"",                // IsReturnFiled (Not provided)
							ret_period         // ReturnPeriod
						);
					}
				}
			}

			return dt;
		}
		
		// Helper method to convert JSON to CSV
		private string ConvertJsonToCsv(string jsonString)
		{
			var csvBuilder = new StringBuilder();
			JArray jsonArray = JArray.Parse(jsonString);

			if (jsonArray.Count == 0)
				throw new Exception("JSON data is empty.");

			// Get columns from the first object
			JObject firstRow = jsonArray[0].Value<JObject>();
			var columns = firstRow.Properties().Select(p => p.Name).ToList();
			csvBuilder.AppendLine(string.Join(",", columns.Select(c => $"\"{c}\"")));

			// Add rows
			foreach (JObject row in jsonArray)
			{
				var fields = columns.Select(c => $"\"{row[c]?.ToString().Replace("\"", "\"\"") ?? ""}\"");
				csvBuilder.AppendLine(string.Join(",", fields));
			}

			return csvBuilder.ToString();
		}

		// Helper method to parse CSV to DataTable
		private DataTable ReadCsvToDataTable(Stream stream)
		{
			DataTable dataTable = new DataTable();
			using (var reader = new StreamReader(stream))
			{
				// Read header
				var headers = reader.ReadLine()?.Split(',').Select(h => h.Trim('\"')).ToArray();
				if (headers == null)
					throw new Exception("CSV data is empty or invalid.");

				foreach (var header in headers)
				{
					dataTable.Columns.Add(header);
				}

				// Read rows
				while (!reader.EndOfStream)
				{
					var row = reader.ReadLine()?.Split(',').Select(f => f.Trim('\"')).ToArray();
					if (row != null && row.Length == headers.Length)
					{
						dataTable.Rows.Add(row);
					}
				}
			}

			return dataTable;
		}


		#endregion

		#region PurchaseRegisterClosedRequests
		public async Task<IActionResult> ClosedRequests(DateTime? fromdate, DateTime? todate)
		{
			// Default to 2 days ago and today if no values are provided
			DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
			DateTime toDateTime = todate ?? DateTime.Now; // Default to today

			//await TjCaptions("CompletedTasks"); // Load captions for CompletedTask page
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act2";
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                return View("~/Views/Admin/ClosedTasks/ClosedTask.cshtml");
            }

            //var Tickets = await _purchaseTicketBusiness.GetAllCloseTicketsStatusAsync(fromDateTime, toDateTime);
            //         var email = MySession.Current.Email; // Get the email from session
            //var clients = await _userBusiness.GetAdminClients(email);
            //string[] gstinArray = clients.Select(c => c.ClientGSTIN).ToArray();
            //var Tickets = await _purchaseTicketBusiness.GetClientsClosedTicketsStatusAsync(gstinArray, fromDateTime, toDateTime);

            var Tickets = await _purchaseTicketBusiness.GetCloseTicketsStatusAsync(MySession.Current.gstin, fromDateTime, toDateTime); //(ViewBag.UserGSTIN); // Fetch pending tickets based on user_gstin

            ViewBag.Tickets = Tickets;

			return View("~/Views/Admin/ClosedTasks/ClosedTask.cshtml");

			// var ticketsJson = HttpContext.Session.GetString("TicketsList");
			//var tickets = JsonConvert.DeserializeObject<List<TicketsStatusModel>>(ticketsJson);
			//string[] clients = { "1", "2", "3" }; // create and call function here
			//step -2
			//Based on this clients Gstin fetch Pending the tickets from Purchase_Ticket table
			//var Tickets = await _purchaseTicketBusiness.GetClientTicketsStatusAsync(clients);
			// HttpContext.Session.SetString("TicketsList", JsonConvert.SerializeObject(Tickets));
		}

		public async Task<IActionResult> CloseRequest(string requestNo, string ClientGSTIN)
		{
			//await TjCaptions("CloseRequest");
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act2"; // For active button styling
			await _purchaseTicketBusiness.UpdatePurchaseTicketAsync(requestNo, ClientGSTIN);

			var ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);

			string toEmail = _configuration["Mail:ToMail"];
			if (_configuration["Mail:SendToClient"] == "Yes")
			{
				toEmail = ticket.Email;
			}
			string fromMail = _configuration["Mail:FromMail"];

			string subjectTemplate = _configuration["Mail:SubjectTemplate"];
			string bodyTemplate = _configuration["Mail:BodyTemplate"];

			string type = _configuration["Mail:Type1"];
			string subject = subjectTemplate
				.Replace("{Type}", type)
				.Replace("{RequestNo}", requestNo);

			Attachment attachment1 = await GenerateInvoiceExcelAttachmentAsync(requestNo, ClientGSTIN);
			Attachment attachment2 = await GenerateExcelAttachmentAsync(requestNo, ticket.FileName, ClientGSTIN);


			string body = bodyTemplate
				.Replace("{CustomerName}", ticket.EXERTUSERNAME)
				.Replace("{RequestNo}", requestNo)
				.Replace("{CreatedDate}", ticket.RequestCreatedDate?.ToString("yyyy-MMM-dd hh:mm:ss tt"))
				.Replace("{ClosedDate}", DateTime.Now.ToString("yyyy-MMM-dd hh:mm:ss tt"))
				.Replace("{FileName}", ticket.FileName)
				.Replace("{OutputFileName}", attachment2.Name);


			// Send email
			string[] ccList = _configuration["Mail:CCMail"].Split(',');

			await SendEmailAsync(
				toEmail,
				subject,
				body,
				ccList,
				attachment1,
				attachment2
			);


			return Json(new { success = true });
		}

		#endregion

		#region Purchase Register Export XL

		public async Task<IActionResult> ExportReport(string requestNo, string fileName, string ClientGSTIN)
		{
			var ticketDetails = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			//var fileName = ticketDetails.FileName;
			var adminFileName = ticketDetails.AdminFileName;
			var status = ticketDetails.TicketStatus;
			fileName = ticketDetails.FileName;
			var clientGSTIN = ticketDetails.ClientGSTIN;
			string ExportFileName = $"{requestNo}_Report.xlsx";
			//ViewBag.Message = "yes";
			if (status == "Completed")
			{
				ExportFileName = $"{fileName.Split('.')[0]}_{requestNo}_Report.xlsx";
			}
			if (status == "Analysed")
			{
				ExportFileName = $"{fileName.Split('.')[0]}_VS_{adminFileName.Split('.')[0]}_Report.xlsx";
			}


			var data = await _compareGstBusiness.GetComparedDataBasedOnTicketAsync(requestNo, clientGSTIN);

			var invoiceData = data
				.Where(d => d.DataSource != null && d.DataSource.Trim().Equals("Invoice", StringComparison.OrdinalIgnoreCase))
				.OrderBy(x => x.RowNumber)
				.ToList();
			//_logger.LogInformation($"Invoice Data Count: {invoiceData.Count}");

			var portalData = data
				.Where(d => d.DataSource != null && d.DataSource.Trim().Equals("Portal", StringComparison.OrdinalIgnoreCase))
				.OrderBy(x => x.RowNumber)
				.ToList();
			//_logger.LogInformation($"Portal Data Count: {portalData.Count}");
			var summary = GenerateSummary(data);


			string[] matchType =
			{
				_configuration["MatchTypes:1"],
				_configuration["MatchTypes:2"],
				_configuration["MatchTypes:3"],
				_configuration["MatchTypes:4"],
				_configuration["MatchTypes:5"],
				_configuration["MatchTypes:6"],
				_configuration["MatchTypes:7"],
				_configuration["MatchTypes:8"],
				_configuration["MatchTypes:9"],
				_configuration["MatchTypes:10"]
			};

			using (var workbook = new XLWorkbook())
			{
				// ✅ Sheet 1: Summary
				var summarySheet = workbook.Worksheets.Add("Summary");
				//summarySheet.Range("A1:D1").Merge().Value = "Merged Cell Value";

				summarySheet.Range("D1:F1").Merge().Value = "SUM";
				summarySheet.Range("G1:I1").Merge().Value = "COUNT";

				summarySheet.Range("D2:F2").Merge().Value = "Total Tax_A";
				summarySheet.Range("G2:I2").Merge().Value = "Total Tax_A";


				summarySheet.Cell(3, 3).Value = "Data Source";

				summarySheet.Cell(4, 1).Value = "Matching Results";
				summarySheet.Cell(4, 2).Value = "Categories";
				summarySheet.Cell(4, 3).Value = "Match Type";

				summarySheet.Range("D3:D4").Merge().Value = "Invoice";
				summarySheet.Range("E3:E4").Merge().Value = "Portal";
				summarySheet.Range("F3:F4").Merge().Value = "Grand Total";
				summarySheet.Range("G3:G4").Merge().Value = "Invoice";
				summarySheet.Range("H3:H4").Merge().Value = "Portal";
				summarySheet.Range("I3:I4").Merge().Value = "Grand Total";
				summarySheet.Range("J3:J4").Merge().Value = "% Matching";


				summarySheet.Range("A1:B3").Merge();
				summarySheet.Range("c1:c2").Merge();

				summarySheet.Range("A1:J4").Style.Font.Bold = true;
				summarySheet.Range("A1:J4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

				var totalinvoicetax = summary.catagory1InvoiceSum + summary.catagory2InvoiceSum + summary.catagory3InvoiceSum +
									  summary.catagory4InvoiceSum + summary.catagory5InvoiceSum + summary.catagory6InvoiceSum +
									  summary.catagory7InvoiceSum + summary.catagory8InvoiceSum;
				// + summary.catagory9InvoiceSum + summary.catagory10InvoiceSum;


				var totalportaltax = summary.catagory1PortalSum + summary.catagory2PortalSum + summary.catagory3PortalSum +
									 summary.catagory4PortalSum + summary.catagory5PortalSum + summary.catagory6PortalSum +
									 summary.catagory7PortalSum + summary.catagory8PortalSum;
				//+ summary.catagory9PortalSum +     summary.catagory10PortalSum;

				var totalinvoicecount = summary.catagory1InvoiceNumber + summary.catagory2InvoiceNumber + summary.catagory3InvoiceNumber +
										summary.catagory4InvoiceNumber + summary.catagory5InvoiceNumber + summary.catagory6InvoiceNumber +
										summary.catagory7InvoiceNumber + summary.catagory8InvoiceNumber;
				//+ summary.catagory9InvoiceNumber +   summary.catagory10InvoiceNumber;

				var totalportalcount = summary.catagory1PortalNumber + summary.catagory2PortalNumber + summary.catagory3PortalNumber +
									   summary.catagory4PortalNumber + summary.catagory5PortalNumber + summary.catagory6PortalNumber +
									   summary.catagory7PortalNumber + summary.catagory8PortalNumber;
				//+ summary.catagory9PortalNumber + summary.catagory10PortalNumber;

				var grandtotaltax = totalinvoicetax - totalportaltax;
				decimal grandtotalcount = (decimal)totalinvoicecount + (decimal)totalportalcount;

				int summaryrow = 5; // Start from the second row for data
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[0])), GetCategory(matchType[0]), matchType[0], summary.catagory1InvoiceSum ?? 0, summary.catagory1PortalSum ?? 0, summary.catagory1InvoiceNumber ?? 0, summary.catagory1PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[1])), GetCategory(matchType[1]), matchType[1], summary.catagory2InvoiceSum ?? 0, summary.catagory2PortalSum ?? 0, summary.catagory2InvoiceNumber ?? 0, summary.catagory2PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[2])), GetCategory(matchType[2]), matchType[2], summary.catagory3InvoiceSum ?? 0, summary.catagory3PortalSum ?? 0, summary.catagory3InvoiceNumber ?? 0, summary.catagory3PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[3])), GetCategory(matchType[3]), matchType[3], summary.catagory4InvoiceSum ?? 0, summary.catagory4PortalSum ?? 0, summary.catagory4InvoiceNumber ?? 0, summary.catagory4PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[4])), GetCategory(matchType[4]), matchType[4], summary.catagory5InvoiceSum ?? 0, summary.catagory5PortalSum ?? 0, summary.catagory5InvoiceNumber ?? 0, summary.catagory5PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[5])), GetCategory(matchType[5]), matchType[5], summary.catagory6InvoiceSum ?? 0, summary.catagory6PortalSum ?? 0, summary.catagory6InvoiceNumber ?? 0, summary.catagory6PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[6])), GetCategory(matchType[6]), matchType[6], summary.catagory7InvoiceSum ?? 0, summary.catagory7PortalSum ?? 0, summary.catagory7InvoiceNumber ?? 0, summary.catagory7PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[7])), GetCategory(matchType[7]), matchType[7], summary.catagory8InvoiceSum ?? 0, summary.catagory8PortalSum ?? 0, summary.catagory8InvoiceNumber ?? 0, summary.catagory8PortalNumber ?? 0, grandtotalcount);
				//AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[8])), GetCategory(matchType[8]), matchType[8], summary.catagory9InvoiceSum ?? 0, summary.catagory9PortalSum ?? 0, summary.catagory9InvoiceNumber ?? 0, summary.catagory9PortalNumber ?? 0);
				//AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[9])), GetCategory(matchType[9]), matchType[9], summary.catagory10InvoiceSum ?? 0, summary.catagory10PortalSum ?? 0, summary.catagory10InvoiceNumber ?? 0, summary.catagory10PortalNumber ?? 0);


				

				summarySheet.Range($"A{summaryrow}:C{summaryrow}").Merge().Value = "Grand Total";
				summarySheet.Range($"A{summaryrow}:C{summaryrow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
				summarySheet.Cell(summaryrow, 4).Value = totalinvoicetax;
				summarySheet.Cell(summaryrow, 5).Value = -1 * totalportaltax;
				summarySheet.Cell(summaryrow, 6).Value = grandtotaltax;
				summarySheet.Cell(summaryrow, 7).Value = totalinvoicecount;
				summarySheet.Cell(summaryrow, 8).Value = totalportalcount;
				summarySheet.Cell(summaryrow, 9).Value = grandtotalcount;
				summarySheet.Cell(summaryrow, 10).Value = 0;

				summarySheet.Range($"A1:J{summaryrow}").Style.Border.RightBorder = XLBorderStyleValues.Thin;
				//summarySheet.Range("A5:J13").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
				summarySheet.Range($"A1:J{summaryrow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
				summarySheet.Range($"A{summaryrow}:J{summaryrow}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;

				summaryrow = summaryrow + 3;
				summarySheet.Cell(summaryrow, 2).Value = "Request Created Date Time";
				summarySheet.Cell(summaryrow, 3).Value = "Request Updated Date Time";
				summarySheet.Cell(summaryrow, 4).Value = "Request Completed Date Time";

				summaryrow++;
				summarySheet.Cell(summaryrow, 2).Value = ticketDetails.RequestCreatedDate;
				summarySheet.Cell(summaryrow, 3).Value = ticketDetails.RequestUpdatedDate;
				summarySheet.Cell(summaryrow, 4).Value = ticketDetails.RequestCompletedDateTime;

				summarySheet.Cell(summaryrow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
				summarySheet.Cell(summaryrow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
				summarySheet.Cell(summaryrow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

				// ✅ Sheet 2: Main Output
				var mainSheet = workbook.Worksheets.Add("Main Output");
				var headers = new[]
				{
					"Sno",
					"User GSTIN",
					"YearMonth",
					"Financial Year",
					"Datasource",
					"Match_Type",
					"Matching_Results",
					"Categories",

					"SupplierGSTIN",
					"ModifiedSupplierGSTIN",

					"SupplierName",

					"InvoiceNo",
					"ModifiedInvoiceNumber",

					"InvoiceDate",
					"ModifiedInvoiceDate",

					"TaxableValue",

					"TotalTax",
					"ModifiedTotalTax",

					"CGST",
					"SGST",
					"IGST",
					"CESS",
					"Period",
				};
				// Write header row
				for (int i = 0; i < headers.Length; i++)
				{
					mainSheet.Cell(1, i + 1).Value = headers[i];
					mainSheet.Cell(1, i + 1).Style.Font.Bold = true;
				}
				int row = 2;
				// ✅ 1. Write Invoice Data (Sno starts from 1)
				int invoiceSno = 1;
				foreach (var item in invoiceData)
				{
					mainSheet.Cell(row, 1).Value = invoiceSno++;
					mainSheet.Cell(row, 2).Value = item.ClientGSTIN;
					mainSheet.Cell(row, 3).Value = item.YearMonth;
					mainSheet.Cell(row, 4).Value = item.FinancialYear;
					mainSheet.Cell(row, 5).Value = item.DataSource;
					mainSheet.Cell(row, 6).Value = item.MatchType;
					mainSheet.Cell(row, 7).Value = item.MatchingResults;
					mainSheet.Cell(row, 8).Value = item.Category;

					mainSheet.Cell(row, 9).Value = item.SupplierGSTIN;
					mainSheet.Cell(row, 10).Value = item.ModifiedSupplierGSTIN;

					mainSheet.Cell(row, 11).Value = item.SupplierName;

					mainSheet.Cell(row, 12).Value = item.InvoiceNumber;
					mainSheet.Cell(row, 13).Value = item.ModifiedInvoiceNumber;

					mainSheet.Cell(row, 14).Value = item.InvoiceDate;
					mainSheet.Cell(row, 15).Value = item.ModifiedInvoiceDate;

					mainSheet.Cell(row, 16).Value = item.TaxableValue;

					mainSheet.Cell(row, 17).Value = item.TotalTax;
					mainSheet.Cell(row, 18).Value = item.ModifiedTotalTax == 0 ? "" : item.ModifiedTotalTax;

					mainSheet.Cell(row, 19).Value = item.CGST;
					mainSheet.Cell(row, 20).Value = item.SGST;
					mainSheet.Cell(row, 21).Value = item.IGST;
					mainSheet.Cell(row, 22).Value = item.CESS;
					mainSheet.Cell(row, 23).Value = item.Period;

					row++;
				}

				// ✅ 2. Write Portal Data (Sno starts from 1 again)
				int portalSno = 1;
				foreach (var item in portalData)
				{
					mainSheet.Cell(row, 1).Value = portalSno++;
					mainSheet.Cell(row, 2).Value = item.ClientGSTIN;
					mainSheet.Cell(row, 3).Value = item.YearMonth;
					mainSheet.Cell(row, 4).Value = item.FinancialYear;
					mainSheet.Cell(row, 5).Value = item.DataSource;
					mainSheet.Cell(row, 6).Value = item.MatchType;
					mainSheet.Cell(row, 7).Value = item.MatchingResults;
					mainSheet.Cell(row, 8).Value = item.Category;

					mainSheet.Cell(row, 9).Value = item.SupplierGSTIN;
					mainSheet.Cell(row, 10).Value = "";

					mainSheet.Cell(row, 11).Value = item.SupplierName;

					mainSheet.Cell(row, 12).Value = item.InvoiceNumber;
					mainSheet.Cell(row, 13).Value = "";

					mainSheet.Cell(row, 14).Value = item.InvoiceDate;
					mainSheet.Cell(row, 15).Value = "";

					mainSheet.Cell(row, 16).Value = item.TaxableValue;

					mainSheet.Cell(row, 17).Value = item.TotalTax;
					mainSheet.Cell(row, 18).Value = "";

					mainSheet.Cell(row, 19).Value = item.CGST;
					mainSheet.Cell(row, 20).Value = item.SGST;
					mainSheet.Cell(row, 21).Value = item.IGST;
					mainSheet.Cell(row, 22).Value = item.CESS;
					mainSheet.Cell(row, 23).Value = item.Period;

					row++;
				}

				// Export Excel
				using (var stream = new MemoryStream())
				{
					workbook.SaveAs(stream);
					var content = stream.ToArray();

					//ViewBag.Message = "yes";

					return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
											$"{ExportFileName}");
				}
			}
			// return View("~/Views/Admin/CompareGstFiles/CompareGST.cshtml"); // Redirect to the same view if needed  
		}

		public async Task<Attachment> GenerateExcelAttachmentAsync(string requestNo, string fileName, string ClientGSTIN)
		{
			var ticketDetails = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			//var fileName = ticketDetails.FileName;
			var adminFileName = ticketDetails.AdminFileName;
			var status = ticketDetails.TicketStatus;
			fileName = ticketDetails.FileName;
			var clientGSTIN = ticketDetails.ClientGSTIN;

			//ViewBag.Message = "yes";
			string ExportFileName = $"{fileName.Split('.')[0]}_{requestNo}_Report.xlsx";



			var data = await _compareGstBusiness.GetComparedDataBasedOnTicketAsync(requestNo, clientGSTIN);

			var invoiceData = data
				.Where(d => d.DataSource != null && d.DataSource.Trim().Equals("Invoice", StringComparison.OrdinalIgnoreCase))
				.OrderBy(x => x.RowNumber)
				.ToList();
			//_logger.LogInformation($"Invoice Data Count: {invoiceData.Count}");

			var portalData = data
				.Where(d => d.DataSource != null && d.DataSource.Trim().Equals("Portal", StringComparison.OrdinalIgnoreCase))
				.OrderBy(x => x.RowNumber)
				.ToList();
			//_logger.LogInformation($"Portal Data Count: {portalData.Count}");
			var summary = GenerateSummary(data);


			string[] matchType =
			{
		  _configuration["MatchTypes:1"],
		  _configuration["MatchTypes:2"],
		  _configuration["MatchTypes:3"],
		  _configuration["MatchTypes:4"],
		  _configuration["MatchTypes:5"],
		  _configuration["MatchTypes:6"],
		  _configuration["MatchTypes:7"],
		  _configuration["MatchTypes:8"],
		  _configuration["MatchTypes:9"],
		  _configuration["MatchTypes:10"]
	  };

			using (var workbook = new XLWorkbook())
			{
				// ✅ Sheet 1: Summary
				var summarySheet = workbook.Worksheets.Add("Summary");
				//summarySheet.Range("A1:D1").Merge().Value = "Merged Cell Value";

				summarySheet.Range("D1:F1").Merge().Value = "SUM";
				summarySheet.Range("G1:I1").Merge().Value = "COUNT";

				summarySheet.Range("D2:F2").Merge().Value = "Total Tax_A";
				summarySheet.Range("G2:I2").Merge().Value = "Total Tax_A";


				summarySheet.Cell(3, 3).Value = "Data Source";

				summarySheet.Cell(4, 1).Value = "Matching Results";
				summarySheet.Cell(4, 2).Value = "Categories";
				summarySheet.Cell(4, 3).Value = "Match Type";

				summarySheet.Range("D3:D4").Merge().Value = "Invoice";
				summarySheet.Range("E3:E4").Merge().Value = "Portal";
				summarySheet.Range("F3:F4").Merge().Value = "Grand Total";
				summarySheet.Range("G3:G4").Merge().Value = "Invoice";
				summarySheet.Range("H3:H4").Merge().Value = "Portal";
				summarySheet.Range("I3:I4").Merge().Value = "Grand Total";
				summarySheet.Range("J3:J4").Merge().Value = "% Matching";


				summarySheet.Range("A1:B3").Merge();
				summarySheet.Range("c1:c2").Merge();

				summarySheet.Range("A1:J4").Style.Font.Bold = true;
				summarySheet.Range("A1:J4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


				var totalinvoicetax = summary.catagory1InvoiceSum + summary.catagory2InvoiceSum + summary.catagory3InvoiceSum +
									  summary.catagory4InvoiceSum + summary.catagory5InvoiceSum + summary.catagory6InvoiceSum +
									  summary.catagory7InvoiceSum + summary.catagory8InvoiceSum;
				// + summary.catagory9InvoiceSum + summary.catagory10InvoiceSum;


				var totalportaltax = summary.catagory1PortalSum + summary.catagory2PortalSum + summary.catagory3PortalSum +
									 summary.catagory4PortalSum + summary.catagory5PortalSum + summary.catagory6PortalSum +
									 summary.catagory7PortalSum + summary.catagory8PortalSum;
				//+ summary.catagory9PortalSum +     summary.catagory10PortalSum;

				var totalinvoicecount = summary.catagory1InvoiceNumber + summary.catagory2InvoiceNumber + summary.catagory3InvoiceNumber +
										summary.catagory4InvoiceNumber + summary.catagory5InvoiceNumber + summary.catagory6InvoiceNumber +
										summary.catagory7InvoiceNumber + summary.catagory8InvoiceNumber;
				//+ summary.catagory9InvoiceNumber +   summary.catagory10InvoiceNumber;

				var totalportalcount = summary.catagory1PortalNumber + summary.catagory2PortalNumber + summary.catagory3PortalNumber +
									   summary.catagory4PortalNumber + summary.catagory5PortalNumber + summary.catagory6PortalNumber +
									   summary.catagory7PortalNumber + summary.catagory8PortalNumber;
				//+ summary.catagory9PortalNumber + summary.catagory10PortalNumber;

				var grandtotaltax = totalinvoicetax - totalportaltax;
				decimal grandtotalcount = (decimal)totalinvoicecount + (decimal)totalportalcount;

				int summaryrow = 5; // Start from the second row for data
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[0])), GetCategory(matchType[0]), matchType[0], summary.catagory1InvoiceSum ?? 0, summary.catagory1PortalSum ?? 0, summary.catagory1InvoiceNumber ?? 0, summary.catagory1PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[1])), GetCategory(matchType[1]), matchType[1], summary.catagory2InvoiceSum ?? 0, summary.catagory2PortalSum ?? 0, summary.catagory2InvoiceNumber ?? 0, summary.catagory2PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[2])), GetCategory(matchType[2]), matchType[2], summary.catagory3InvoiceSum ?? 0, summary.catagory3PortalSum ?? 0, summary.catagory3InvoiceNumber ?? 0, summary.catagory3PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[3])), GetCategory(matchType[3]), matchType[3], summary.catagory4InvoiceSum ?? 0, summary.catagory4PortalSum ?? 0, summary.catagory4InvoiceNumber ?? 0, summary.catagory4PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[4])), GetCategory(matchType[4]), matchType[4], summary.catagory5InvoiceSum ?? 0, summary.catagory5PortalSum ?? 0, summary.catagory5InvoiceNumber ?? 0, summary.catagory5PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[5])), GetCategory(matchType[5]), matchType[5], summary.catagory6InvoiceSum ?? 0, summary.catagory6PortalSum ?? 0, summary.catagory6InvoiceNumber ?? 0, summary.catagory6PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[6])), GetCategory(matchType[6]), matchType[6], summary.catagory7InvoiceSum ?? 0, summary.catagory7PortalSum ?? 0, summary.catagory7InvoiceNumber ?? 0, summary.catagory7PortalNumber ?? 0, grandtotalcount);
				AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[7])), GetCategory(matchType[7]), matchType[7], summary.catagory8InvoiceSum ?? 0, summary.catagory8PortalSum ?? 0, summary.catagory8InvoiceNumber ?? 0, summary.catagory8PortalNumber ?? 0, grandtotalcount);
				//AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[8])), GetCategory(matchType[8]), matchType[8], summary.catagory9InvoiceSum ?? 0, summary.catagory9PortalSum ?? 0, summary.catagory9InvoiceNumber ?? 0, summary.catagory9PortalNumber ?? 0);
				//AddRow(summarySheet, summaryrow++, GetMatching_Results(GetCategory(matchType[9])), GetCategory(matchType[9]), matchType[9], summary.catagory10InvoiceSum ?? 0, summary.catagory10PortalSum ?? 0, summary.catagory10InvoiceNumber ?? 0, summary.catagory10PortalNumber ?? 0);



				summarySheet.Range($"A{summaryrow}:C{summaryrow}").Merge().Value = "Grand Total";
				summarySheet.Range($"A{summaryrow}:C{summaryrow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
				summarySheet.Cell(summaryrow, 4).Value = totalinvoicetax;
				summarySheet.Cell(summaryrow, 5).Value = -1 * totalportaltax;
				summarySheet.Cell(summaryrow, 6).Value = grandtotaltax;
				summarySheet.Cell(summaryrow, 7).Value = totalinvoicecount;
				summarySheet.Cell(summaryrow, 8).Value = totalportalcount;
				summarySheet.Cell(summaryrow, 9).Value = grandtotalcount;
				summarySheet.Cell(summaryrow, 10).Value = 0;

				summarySheet.Range($"A1:J{summaryrow}").Style.Border.RightBorder = XLBorderStyleValues.Thin;
				//summarySheet.Range("A5:J13").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
				summarySheet.Range($"A1:J{summaryrow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
				summarySheet.Range($"A{summaryrow}:J{summaryrow}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;

				summaryrow = summaryrow + 3;
				summarySheet.Cell(summaryrow, 2).Value = "Request Created Date Time";
				summarySheet.Cell(summaryrow, 3).Value = "Request Updated Date Time";
				summarySheet.Cell(summaryrow, 4).Value = "Request Completed Date Time";

				summaryrow++;
				summarySheet.Cell(summaryrow, 2).Value = ticketDetails.RequestCreatedDate;
				summarySheet.Cell(summaryrow, 3).Value = ticketDetails.RequestUpdatedDate;
				summarySheet.Cell(summaryrow, 4).Value = ticketDetails.RequestCompletedDateTime;

				summarySheet.Cell(summaryrow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
				summarySheet.Cell(summaryrow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
				summarySheet.Cell(summaryrow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

				// ✅ Sheet 2: Main Output
				var mainSheet = workbook.Worksheets.Add("Main Output");
				var headers = new[]
				{
			  "Sno",
			  "User GSTIN",
			  "YearMonth",
			  "Financial Year",
			  "Datasource",
			  "Match_Type",
			  "Matching_Results",
			  "Categories",

			  "SupplierGSTIN",
			  "ModifiedSupplierGSTIN",

			  "SupplierName",

			  "InvoiceNo",
			  "ModifiedInvoiceNumber",

			  "InvoiceDate",
			  "ModifiedInvoiceDate",

			  "TaxableValue",

			  "TotalTax",
			  "ModifiedTotalTax",

			  "CGST",
			  "SGST",
			  "IGST",
			  "CESS",
			  "Period",
		  };
				// Write header row
				for (int i = 0; i < headers.Length; i++)
				{
					mainSheet.Cell(1, i + 1).Value = headers[i];
					mainSheet.Cell(1, i + 1).Style.Font.Bold = true;
				}
				int row = 2;
				// ✅ 1. Write Invoice Data (Sno starts from 1)
				int invoiceSno = 1;
				foreach (var item in invoiceData)
				{
					mainSheet.Cell(row, 1).Value = invoiceSno++;
					mainSheet.Cell(row, 2).Value = item.ClientGSTIN;
					mainSheet.Cell(row, 3).Value = item.YearMonth;
					mainSheet.Cell(row, 4).Value = item.FinancialYear;
					mainSheet.Cell(row, 5).Value = item.DataSource;
					mainSheet.Cell(row, 6).Value = item.MatchType;
					mainSheet.Cell(row, 7).Value = item.MatchingResults;
					mainSheet.Cell(row, 8).Value = item.Category;

					mainSheet.Cell(row, 9).Value = item.SupplierGSTIN;
					mainSheet.Cell(row, 10).Value = item.ModifiedSupplierGSTIN;

					mainSheet.Cell(row, 11).Value = item.SupplierName;

					mainSheet.Cell(row, 12).Value = item.InvoiceNumber;
					mainSheet.Cell(row, 13).Value = item.ModifiedInvoiceNumber;

					mainSheet.Cell(row, 14).Value = item.InvoiceDate;
					mainSheet.Cell(row, 15).Value = item.ModifiedInvoiceDate;

					mainSheet.Cell(row, 16).Value = item.TaxableValue;

					mainSheet.Cell(row, 17).Value = item.TotalTax;
					mainSheet.Cell(row, 18).Value = item.ModifiedTotalTax == 0 ? "" : item.ModifiedTotalTax;

					mainSheet.Cell(row, 19).Value = item.CGST;
					mainSheet.Cell(row, 20).Value = item.SGST;
					mainSheet.Cell(row, 21).Value = item.IGST;
					mainSheet.Cell(row, 22).Value = item.CESS;
					mainSheet.Cell(row, 23).Value = item.Period;

					row++;
				}

				// ✅ 2. Write Portal Data (Sno starts from 1 again)
				int portalSno = 1;
				foreach (var item in portalData)
				{
					mainSheet.Cell(row, 1).Value = portalSno++;
					mainSheet.Cell(row, 2).Value = item.ClientGSTIN;
					mainSheet.Cell(row, 3).Value = item.YearMonth;
					mainSheet.Cell(row, 4).Value = item.FinancialYear;
					mainSheet.Cell(row, 5).Value = item.DataSource;
					mainSheet.Cell(row, 6).Value = item.MatchType;
					mainSheet.Cell(row, 7).Value = item.MatchingResults;
					mainSheet.Cell(row, 8).Value = item.Category;

					mainSheet.Cell(row, 9).Value = item.SupplierGSTIN;
					mainSheet.Cell(row, 10).Value = "";

					mainSheet.Cell(row, 11).Value = item.SupplierName;

					mainSheet.Cell(row, 12).Value = item.InvoiceNumber;
					mainSheet.Cell(row, 13).Value = "";

					mainSheet.Cell(row, 14).Value = item.InvoiceDate;
					mainSheet.Cell(row, 15).Value = "";

					mainSheet.Cell(row, 16).Value = item.TaxableValue;

					mainSheet.Cell(row, 17).Value = item.TotalTax;
					mainSheet.Cell(row, 18).Value = "";

					mainSheet.Cell(row, 19).Value = item.CGST;
					mainSheet.Cell(row, 20).Value = item.SGST;
					mainSheet.Cell(row, 21).Value = item.IGST;
					mainSheet.Cell(row, 22).Value = item.CESS;
					mainSheet.Cell(row, 23).Value = item.Period;

					row++;
				}

				var stream = new MemoryStream();

				workbook.SaveAs(stream);
				stream.Position = 0; // Reset stream before use

				var attachment = new Attachment(stream, ExportFileName,
					"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

				return attachment;

			}
		}

		public async Task<IActionResult> ExportInvoiceFile(string requestNo, string ClientGSTIN)
		{
			var ticketDetails = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			var data = await _purchaseDataBusiness.GetInvoiceData(requestNo , ClientGSTIN);
			
			var fileName = ticketDetails.FileName;
			string[] parts = fileName.Split('.');
			string fileExtension = parts[1];
		   
			var _invoice = _configuration["Invoice"];
			string[] headers = _invoice.Split(',').Select(x => x.Trim()).ToArray();

			if (fileExtension == "csv")
			{
				var csvBuilder = new StringBuilder();
				// Write headers
				csvBuilder.AppendLine(string.Join(",", headers));
				// Write data
				foreach (var item in data)
				{
					var row = new string[]
					{
						item.ClientGSTIN,
						item.TxnPeriod,
						item.SupplierGSTIN,
						item.SupplierName,
						item.Invoiceno,
						item.InvoiceDate?.ToString("yyyy-MM-dd") ?? "", // format date
						item.TaxableValue.ToString(),
						item.CGST.ToString(),
						item.SGST.ToString(),
						item.IntegratedTax.ToString(),
						item.CESS.ToString()
					};
				   csvBuilder.AppendLine(string.Join(",", row.Select(field => EscapeCsv(field))));
				}
				byte[] content = Encoding.UTF8.GetBytes(csvBuilder.ToString());
				return File(content, "text/csv", $"{fileName}");
			}
			if (fileExtension == "xlsx")
			{
				using var workbook = new XLWorkbook();
				var worksheet = workbook.Worksheets.Add("Purchase Register");
				for (int i = 0; i < headers.Length; i++)
				{
					worksheet.Cell(1, i + 1).Value = headers[i];
					worksheet.Cell(1, i + 1).Style.Font.Bold = true;
				}
				int rowIndex = 2;
				foreach (var item in data)
				{
					worksheet.Cell(rowIndex, 1).Value = item.ClientGSTIN;
					worksheet.Cell(rowIndex, 2).Value = item.TxnPeriod;
					worksheet.Cell(rowIndex, 3).Value = item.SupplierGSTIN;
					worksheet.Cell(rowIndex, 4).Value = item.SupplierName;
					worksheet.Cell(rowIndex, 5).Value = item.Invoiceno;
					worksheet.Cell(rowIndex, 6).Value = item.InvoiceDate?.ToString("yyyy-MM-dd") ?? ""; // format date
					worksheet.Cell(rowIndex, 7).Value = item.TaxableValue.ToString();
					worksheet.Cell(rowIndex, 8).Value = item.CGST.ToString();
					worksheet.Cell(rowIndex, 9).Value = item.SGST.ToString();
					worksheet.Cell(rowIndex, 10).Value = item.IntegratedTax.ToString();
					worksheet.Cell(rowIndex, 11).Value = item.CESS.ToString();

					rowIndex++;
				}
				var stream = new MemoryStream();
				workbook.SaveAs(stream);
				stream.Position = 0; // Reset stream before using
				return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{fileName}");
			}
			if (fileExtension == "xls")
			{
				HSSFWorkbook workbook = new HSSFWorkbook();
				ISheet sheet = workbook.CreateSheet("Purchase Register");

				var headerRow = sheet.CreateRow(0);
				for (int i = 0; i < headers.Length; i++)
				{
					headerRow.CreateCell(i).SetCellValue(headers[i]);
				}

				int rowIndex = 1;
				foreach (var item in data)
				{
					var row = sheet.CreateRow(rowIndex++);
					row.CreateCell(0).SetCellValue(item.ClientGSTIN);
					row.CreateCell(1).SetCellValue(item.TxnPeriod);
					row.CreateCell(2).SetCellValue(item.SupplierGSTIN);
					row.CreateCell(3).SetCellValue(item.SupplierName);
					row.CreateCell(4).SetCellValue(item.Invoiceno);
					row.CreateCell(5).SetCellValue(item.InvoiceDate?.ToString("yyyy-MM-dd") ?? ""); // format date
					row.CreateCell(6).SetCellValue(item.TaxableValue.ToString());
					row.CreateCell(7).SetCellValue(item.CGST.ToString());
					row.CreateCell(8).SetCellValue(item.SGST.ToString());
					row.CreateCell(9).SetCellValue(item.IntegratedTax.ToString());
					row.CreateCell(10).SetCellValue(item.CESS.ToString());
				}
				var stream = new MemoryStream();
				workbook.Write(stream);
				stream.Position = 0; // Reset stream before using
				return File(stream.ToArray(), "application/vnd.ms-excel", $"{fileName}");

			}

			return BadRequest("Unsupported file format.");
		}

		public async Task<Attachment> GenerateInvoiceExcelAttachmentAsync(string requestNo, string ClientGSTIN)
		{
			var ticketDetails = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			var data = await _purchaseDataBusiness.GetInvoiceData(requestNo , ClientGSTIN);

			var fileName = ticketDetails.FileName;
			string[] parts = fileName.Split('.');
			string fileExtension = parts[1];

			var _invoice = _configuration["Invoice"];
			string[] headers = _invoice.Split(',').Select(x => x.Trim()).ToArray();

			Attachment attachment = null!;

			if (fileExtension == "csv")
			{
				var csvBuilder = new StringBuilder();
				// Write headers
				csvBuilder.AppendLine(string.Join(",", headers));
				// Write data
				foreach (var item in data)
				{
					var row = new string[]
					{
						  item.ClientGSTIN,
						  item.TxnPeriod,
						  item.SupplierGSTIN,
						  item.SupplierName,
						  item.Invoiceno,
						  item.InvoiceDate?.ToString("yyyy-MM-dd") ?? "", // format date
						  item.TaxableValue.ToString(),
						  item.CGST.ToString(),
						  item.SGST.ToString(),
						  item.IntegratedTax.ToString(),
						  item.CESS.ToString()
					};
					csvBuilder.AppendLine(string.Join(",", row.Select(field => EscapeCsv(field))));
				}
				byte[] content = Encoding.UTF8.GetBytes(csvBuilder.ToString());
				var stream = new MemoryStream(content);
				stream.Position = 0; // Reset stream before using
				attachment = new Attachment(stream, fileName, "text/csv");
			}
			if (fileExtension == "xlsx")
			{
				using var workbook = new XLWorkbook();
				var worksheet = workbook.Worksheets.Add("PR Invoices");
				for (int i = 0; i < headers.Length; i++)
				{
					worksheet.Cell(1, i + 1).Value = headers[i];
					worksheet.Cell(1, i + 1).Style.Font.Bold = true;
				}
				int rowIndex = 2;
				foreach (var item in data)
				{
					worksheet.Cell(rowIndex, 1).Value = item.ClientGSTIN;
					worksheet.Cell(rowIndex, 2).Value = item.TxnPeriod;
					worksheet.Cell(rowIndex, 3).Value = item.SupplierGSTIN;
					worksheet.Cell(rowIndex, 4).Value = item.SupplierName;
					worksheet.Cell(rowIndex, 5).Value = item.Invoiceno;
					worksheet.Cell(rowIndex, 6).Value = item.InvoiceDate?.ToString("yyyy-MM-dd") ?? ""; // format date
					worksheet.Cell(rowIndex, 7).Value = item.TaxableValue.ToString();
					worksheet.Cell(rowIndex, 8).Value = item.CGST.ToString();
					worksheet.Cell(rowIndex, 9).Value = item.SGST.ToString();
					worksheet.Cell(rowIndex, 10).Value = item.IntegratedTax.ToString();
					worksheet.Cell(rowIndex, 11).Value = item.CESS.ToString();
				}
				var stream = new MemoryStream();
				workbook.SaveAs(stream);
				stream.Position = 0; // Reset stream before using
				//return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{fileName}");
				attachment = new Attachment(stream, fileName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
			}
			if (fileExtension == "xls")
			{
				HSSFWorkbook workbook = new HSSFWorkbook();
				ISheet sheet = workbook.CreateSheet("SL Invoices");

				var headerRow = sheet.CreateRow(0);
				for (int i = 0; i < headers.Length; i++)
				{
					headerRow.CreateCell(i).SetCellValue(headers[i]);
				}

				int rowIndex = 1;
				foreach (var item in data)
				{
					var row = sheet.CreateRow(rowIndex++);
					row.CreateCell(0).SetCellValue(item.ClientGSTIN);
					row.CreateCell(1).SetCellValue(item.TxnPeriod);
					row.CreateCell(2).SetCellValue(item.SupplierGSTIN);
					row.CreateCell(3).SetCellValue(item.SupplierName);
					row.CreateCell(4).SetCellValue(item.Invoiceno);
					row.CreateCell(5).SetCellValue(item.InvoiceDate?.ToString("yyyy-MM-dd") ?? ""); // format date
					row.CreateCell(6).SetCellValue(item.TaxableValue.ToString());
					row.CreateCell(7).SetCellValue(item.CGST.ToString());
					row.CreateCell(8).SetCellValue(item.SGST.ToString());
					row.CreateCell(9).SetCellValue(item.IntegratedTax.ToString());
					row.CreateCell(10).SetCellValue(item.CESS.ToString());
				}
				var stream = new MemoryStream();
				workbook.Write(stream);
				stream.Position = 0; // Reset stream before using
				//return File(stream.ToArray(), "application/vnd.ms-excel", $"{fileName}");
				attachment = new Attachment(stream, fileName, "application/vnd.ms-excel");
			}

			return attachment;
		}
	  
		public async Task<IActionResult> ExportPortalFile(string requestNo, string ClientGSTIN)
		{
			var ticketDetails = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			var data = await _gSTR2DataBusiness.GetPortalData(requestNo , ClientGSTIN);
			
			var fileName = ticketDetails.AdminFileName;
			string[] parts = fileName.Split('.');
			string fileExtension = parts[1];

			var _portal = _configuration["Portal"];
			string[] headers = _portal.Split(',').Select(x => x.Trim()).ToArray();

			if (fileExtension == "csv")
			{
				var csvBuilder = new StringBuilder();
				// Write headers
				csvBuilder.AppendLine(string.Join(",", headers));
				// Write data
				foreach (var item in data)
				{
					var row = new string[]
					{
						item.ClientGSTIN,
						item.GSTRegType,
						item.SupplierInvoiceNo,
						item.SupplierInvoiceDate?.ToString("yyyy-MM-dd") ?? "", // format date
						item.SupplierGSTIN,
						item.SupplierName,
						item.IsRCMApplied?.ToString() ?? "",
						item.InvoiceValue.ToString(),
						item.ItemTaxableValue.ToString(),
						item.GSTRate.ToString(),
						item.IGSTValue.ToString(),
						item.CGSTValue.ToString(),
						item.SGSTValue.ToString(),
						item.CESS.ToString(),
						item.IsReturnFiled?.ToString() ?? "",
						item.TxnPeriod
					};
					csvBuilder.AppendLine(string.Join(",", row.Select(field => EscapeCsv(field))));
				}
				byte[] content = Encoding.UTF8.GetBytes(csvBuilder.ToString());
				return File(content, "text/csv", $"{fileName}");
			}
			if (fileExtension == "xlsx")
			{
				using var workbook = new XLWorkbook();
				var worksheet = workbook.Worksheets.Add("Purchase Register");
				for (int i = 0; i < headers.Length; i++)
				{
					worksheet.Cell(1, i + 1).Value = headers[i];
					worksheet.Cell(1, i + 1).Style.Font.Bold = true;
				}
				int rowIndex = 2;
				foreach (var item in data)
				{
					worksheet.Cell(rowIndex, 1).Value = item.ClientGSTIN;
					worksheet.Cell(rowIndex, 2).Value = item.GSTRegType;
					worksheet.Cell(rowIndex, 3).Value = item.SupplierInvoiceNo;
					worksheet.Cell(rowIndex, 4).Value = item.SupplierInvoiceDate?.ToString("yyyy-MM-dd") ?? ""; // format date
					worksheet.Cell(rowIndex, 5).Value = item.SupplierGSTIN;
					worksheet.Cell(rowIndex, 6).Value = item.SupplierName;
					worksheet.Cell(rowIndex, 7).Value = item.IsRCMApplied?.ToString() ?? "";
					worksheet.Cell(rowIndex, 8).Value = item.InvoiceValue.ToString();
					worksheet.Cell(rowIndex, 9).Value = item.ItemTaxableValue.ToString();
					worksheet.Cell(rowIndex, 10).Value = item.GSTRate.ToString();
					worksheet.Cell(rowIndex, 11).Value = item.IGSTValue.ToString();
					worksheet.Cell(rowIndex, 12).Value = item.CGSTValue.ToString();
					worksheet.Cell(rowIndex, 13).Value = item.SGSTValue.ToString();
					worksheet.Cell(rowIndex, 14).Value = item.CESS.ToString();
					worksheet.Cell(rowIndex, 15).Value = item.IsReturnFiled?.ToString() ?? "";
					worksheet.Cell(rowIndex, 16).Value = item.TxnPeriod;

					rowIndex++; 
				}
				var stream = new MemoryStream();
				workbook.SaveAs(stream);
				stream.Position = 0; // Reset stream before using
				return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{fileName}");                
			}
			if (fileExtension == "xls")
			{
				HSSFWorkbook workbook = new HSSFWorkbook();
				ISheet sheet = workbook.CreateSheet("Purchase Register");


                var headerRow = sheet.CreateRow(0);
				for (int i = 0; i < headers.Length; i++)
				{
					headerRow.CreateCell(i).SetCellValue(headers[i]);
				}

				int rowIndex = 1;
				foreach (var item in data)
				{
					var row = sheet.CreateRow(rowIndex++);
					row.CreateCell(0).SetCellValue(item.ClientGSTIN);
					row.CreateCell(1).SetCellValue(item.GSTRegType);
					row.CreateCell(2).SetCellValue(item.SupplierInvoiceNo);
					row.CreateCell(3).SetCellValue(item.SupplierInvoiceDate?.ToString("yyyy-MM-dd") ?? ""); // format date
					row.CreateCell(4).SetCellValue(item.SupplierGSTIN);
					row.CreateCell(5).SetCellValue(item.SupplierName);
					row.CreateCell(6).SetCellValue(item.IsRCMApplied?.ToString() ?? "");
					row.CreateCell(7).SetCellValue(item.InvoiceValue.ToString());
					row.CreateCell(8).SetCellValue(item.ItemTaxableValue.ToString());
					row.CreateCell(9).SetCellValue(item.GSTRate.ToString());
					row.CreateCell(10).SetCellValue(item.IGSTValue.ToString());
					row.CreateCell(11).SetCellValue(item.CGSTValue.ToString());
					row.CreateCell(12).SetCellValue(item.SGSTValue.ToString());
					row.CreateCell(13).SetCellValue(item.CESS.ToString());
					row.CreateCell(14).SetCellValue(item.IsReturnFiled?.ToString() ?? "");
					row.CreateCell(15).SetCellValue(item.TxnPeriod);                
				}
				var stream = new MemoryStream();
				workbook.Write(stream);
				stream.Position = 0; // Reset stream before using
				return File(stream.ToArray(), "application/vnd.ms-excel", $"{fileName}");                              
			}
			return BadRequest("Unsupported file format.");
		}

		// Helper to escape commas and quotes
		private string EscapeCsv(string field)
		{
			if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
			{
				return $"\"{field.Replace("\"", "\"\"")}\"";
			}
			return field;
		}

		private string GetCategory(string Match_Type)
		{
			switch (Match_Type)
			{
				case "6_UnMatched_Excess_or_Short_1_Invoicewise":
				case "Available_In_PR_Not_In_Portal":
				case "Available_In_Portal_Not_In_PR":
					return "UnMatched";

				case "5_Probable_Matched_GST_TAXB":
				case "4_Partially_Matched_GST_DT":
				case "3_Partially_Matched_GST_INV":
					return "Partially_Matched";

				case "1_Exactly_Matched_GST_INV_DT_TAXB_TAX":
				case "2_Matched_With_Tolerance_GST_INV_DT_TAXB_TAX":
					return "Completely_Matched";

				default:
					return "Unknown";
			}
		}
		private string GetMatching_Results(string Category)
		{
			switch (Category)
			{
				case "UnMatched":
					return "IMS Pending";

				case "Partially_Matched":
				case "Completely_Matched":
					return "IMS Accept";
				default:
					return "Unknown";
			}

		}

		public SummaryExcelModel GenerateSummary(List<CompareGstResultModel> data)
		{
			//_logger.LogInformation("Data count : {0}", data.Count);


			var summary = new SummaryExcelModel();

			// Dictionary of category names mapped to numbers (to match your model)
			var categoryMap = new Dictionary<string, int>
			{
				{ "1_Exactly_Matched_GST_INV_DT_TAXB_TAX", 1 },
				{ "2_Matched_With_Tolerance_GST_INV_DT_TAXB_TAX", 2 },

				{ "3_Partially_Matched_GST_INV", 3 },
				{ "4_Partially_Matched_GST_DT", 4 },
				{ "5_Probable_Matched_GST_TAXB", 5 },

				{ "6_UnMatched_Excess_or_Short_1_Invoicewise", 6 },
				{ "Available_In_PR_Not_In_Portal", 7 },
				{ "Available_In_Portal_Not_In_PR", 8 }
			};
			//_logger.LogInformation("Distinct categories in data: " + string.Join(", ", data.Select(d => d.MatchType).Distinct()));

			foreach (var kvp in categoryMap)
			{
				var categoryName = kvp.Key;
				var categoryNumber = kvp.Value;

				//var categoryData = data.Where(d => d.MatchType == categoryName);
				var categoryData = data.Where(d => d.MatchType?.Trim().Equals(categoryName, StringComparison.OrdinalIgnoreCase) == true);
				//_logger.LogInformation($"Category {categoryNumber} Count: {categoryData.Count()}");

				//_logger.LogInformation("DataSource Values: " + string.Join(", ", categoryData.Select(d => d.DataSource).Distinct()));


				var portalGroup = categoryData.Where(d => d.DataSource != null && d.DataSource.ToLower().Trim() == "portal");
				//_logger.LogInformation($"portalGroup Count: {portalGroup.Count()}");

				var invoiceGroup = categoryData.Where(d => d.DataSource != null && d.DataSource.ToLower().Trim() == "invoice");
				//_logger.LogInformation($"invoiceGroup Count: {invoiceGroup.Count()}");


				typeof(SummaryExcelModel).GetProperty($"catagory{categoryNumber}PortalNumber")?.SetValue(summary, (decimal)portalGroup.Count());
				typeof(SummaryExcelModel).GetProperty($"catagory{categoryNumber}PortalSum")?.SetValue(summary, portalGroup.Sum(x => (decimal?)x.TotalTax) ?? 0);

				typeof(SummaryExcelModel).GetProperty($"catagory{categoryNumber}InvoiceNumber")?.SetValue(summary, (decimal)invoiceGroup.Count());
				typeof(SummaryExcelModel).GetProperty($"catagory{categoryNumber}InvoiceSum")?.SetValue(summary, invoiceGroup.Sum(x => (decimal?)x.TotalTax) ?? 0);


			}

			return summary;
		}

		private void AddRow(IXLWorksheet sheet, int row, string matchingResult, string category, string matchType,
		decimal invoiceSum, decimal portalSum, decimal invoicecount, decimal portalcount, decimal grandtotalcount)
		{
			sheet.Cell(row, 1).Value = matchingResult;
			sheet.Cell(row, 2).Value = category;
			sheet.Cell(row, 3).Value = matchType;
			sheet.Cell(row, 4).Value = invoiceSum;
			sheet.Cell(row, 5).Value = -1 * portalSum;
			sheet.Cell(row, 6).Value = invoiceSum - portalSum;
			sheet.Cell(row, 7).Value = invoicecount;
			sheet.Cell(row, 8).Value = portalcount;
			sheet.Cell(row, 9).Value = invoicecount + portalcount;
			sheet.Cell(row, 10).Value = $"{Math.Round(((invoicecount + portalcount) / (decimal)grandtotalcount) * 100, 2)}%";

		}

		#endregion    
	  
		#region Function to redirect to Purchase Register CompareCSVResults for Admin
		public async Task<IActionResult> CompareCSVResults(string requestNo, string ClientGSTIN)
		{
			ViewBag.Messages = "Admin";
			// Console.WriteLine($"ticketno" + requestNo);
			var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			ViewBag.Ticket = Ticket;
			var clientGSTIN = Ticket.ClientGSTIN;
			var data = await _compareGstBusiness.GetComparedDataBasedOnTicketAsync(requestNo, clientGSTIN);
			ViewBag.ReportDataList = data;
			// _logger.LogInformation("Total Compared Data Rows: " + data.Count);

			// store number and sum in a model and save it in 
			ViewBag.Summary = GenerateSummary(data);
			var summary = ViewBag.Summary;
			//_logger.LogInformation($"invoice 1 count {summary.catagory1InvoiceNumber},portal 1 count {summary.catagory1PortalNumber}");

			return View("/Views/Admin/CompareGstFiles/Compare.cshtml");
		}

		#endregion 

		#region Download Sample Purchase Register Portal File
		[HttpGet]
		public IActionResult DownloadSampleFileCSV()
		{
			string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "PRSamplePortal.csv");

			if (!System.IO.File.Exists(filePath))
			{
				return NotFound("Sample file not found.");
			}

			var fileBytes = System.IO.File.ReadAllBytes(filePath);
			return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PRSamplePortal.csv");
		}
		[HttpGet]
		public IActionResult DownloadSampleFileXLS()
		{
			string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "PRSamplePortal.xls");

			if (!System.IO.File.Exists(filePath))
			{
				return NotFound("Sample file not found.");
			}

			var fileBytes = System.IO.File.ReadAllBytes(filePath);
			return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PRSamplePortal.xls");
		}
		[HttpGet]
		public IActionResult DownloadSampleFileXLSX()
		{
			string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "PRSamplePortal.xlsx");

			if (!System.IO.File.Exists(filePath))
			{
				return NotFound("Sample file not found.");
			}

			var fileBytes = System.IO.File.ReadAllBytes(filePath);
			return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PRSamplePortal.xlsx");
		}

		#endregion

		#region Change Password
		public IActionResult ChangePassword()
		{
			ViewBag.Messages = "Admin";
			return View("~/Views/Admin/ChangePassword/ChangePassword.cshtml");
		}

		[HttpPost]
		public async Task<IActionResult> ChangePassword(string currentPassword, string newPassword, string confirmPassword)
		{
			ViewBag.Messages = "Admin";
            string emailId = MySession.Current.Email;

            var requestBody1 = new { LanguageId = "EN", EmailId = emailId, Password = currentPassword };
            var content1 = new StringContent(JsonConvert.SerializeObject(requestBody1), Encoding.UTF8, "application/json");
            try
            {
                string baseUrl1 = _configuration["ApiSettings:BaseUrl"] ?? throw new InvalidOperationException("BaseUrl not configured");
                string validateLoginEndpoint = _configuration["ApiSettings:ValidateLoginEndpoint"] ?? throw new InvalidOperationException("ValidateLoginEndpoint not configured");
                string apiUrl1 = $"{baseUrl1}{validateLoginEndpoint}";

                var httpClient1 = _httpClientFactory.CreateClient();
                var response = await httpClient1.PostAsync(apiUrl1, content1);
                var responseData = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                }
                else
                {
                    ViewBag.Message = "Invalid Current Password.Please Check...";
                    return View("~/Views/Admin/ChangePassword/ChangePassword.cshtml");
                }
            }
            catch (Exception ex)
            {
                // Handle unexpected exceptions (e.g., network error, timeout)
                ViewBag.ErrorMessage = $"An error occurred while changing password: {ex.Message}";
                return View("~/Views/Admin/ChangePassword/ChangePassword.cshtml");
            }

            var httpClient = _httpClientFactory.CreateClient();
            var requestBody = new
            {
                LanguageId = "EN",
                EmailId = emailId,
                OldPassword = currentPassword,
                NewPassword = newPassword,
            };

            string jsonBody = JsonConvert.SerializeObject(requestBody);
            string baseUrl = _configuration["ApiSettings:BaseUrl"] ?? throw new InvalidOperationException("ApiSettings:BaseUrl is not configured");
            string ChangePasswordEndpoint = _configuration["ApiSettings:ChangePasswordEndpoint"] ?? throw new InvalidOperationException("ApiSettings:ChangePasswordEndpoint is not configured");
            string apiUrl = $"{baseUrl}{ChangePasswordEndpoint}";
            var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            try
            {
                var response = await httpClient.PostAsync(apiUrl, content);
                var responseData = await response.Content.ReadAsStringAsync();
                Console.WriteLine(responseData);

                if (response.IsSuccessStatusCode)
                {
                    // Parse JSON to dynamic or custom model
                    var responseObject = JsonConvert.DeserializeObject<dynamic>(responseData);

                    // Safely extract the return message
                    string returnMessage = responseObject?.returnMessage ?? "Password changed successfully.";

                    ViewBag.Message = returnMessage;
                    return View("~/Views/Admin/ChangePassword/ChangePassword.cshtml");
                }
                else
                {
                    // Handle failure (non-2xx response)
                    var responseObject = JsonConvert.DeserializeObject<dynamic>(responseData);
                    var errorMessage = responseObject?.errorMessage;
                    ViewBag.Message = errorMessage;
                    return View("~/Views/Admin/ChangePassword/ChangePassword.cshtml");
                }
            }
            catch (Exception ex)
            {
                // Handle unexpected exceptions (e.g., network error, timeout)
                ViewBag.ErrorMessage = $"An error occurred while changing password: {ex.Message}";
                return View("~/Views/Admin/ChangePassword/ChangePassword.cshtml");
            }
        }

		#endregion

		#region Logout
		// ✅ Logout (Redirect to Login Page)
		public IActionResult Logout()
		{
			HttpContext.Session.Clear();
			// Set cache headers
			Response.Headers["Cache-Control"] = "no-cache, no-store, must-revalidate, private";
			Response.Headers["Pragma"] = "no-cache";
			Response.Headers["Expires"] = "0";
			return RedirectToAction("Index", "Home");
		}

		#endregion

		#region SendEmail
		public async Task SendEmailAsync(string toEmail, string subject, string body, string[] ccEmails, Attachment attachment1, Attachment attachment2)
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

				// Attachments
				if (attachment1 != null)
				{
					mail.Attachments.Add(attachment1);
				}
				if (attachment2 != null)
				{
					mail.Attachments.Add(attachment2);
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

		#region ClientMaster
		public IActionResult ClientMaster()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act1"; // For active button styling
			return View("~/Views/Future/ComingSoon.cshtml");
		}

        #endregion

        #region SalesLedgerUploadInvoiceFile
        public async Task<IActionResult> SalesLedgerUploadFile(string requestNo)
        {
            var clientGSTIN = MySession.Current.gstin;
            var ticketDetails = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(requestNo, clientGSTIN);
            ViewBag.RequestNo = requestNo;
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            ViewBag.FinancialYear = ticketDetails.FinancialYear;
            ViewBag.PeriodType = ticketDetails.PeriodType;
            ViewBag.TxnPeriod = ticketDetails.Period;
            ViewBag.FileName = ticketDetails.FileName;
            return View("~/Views/Admin/SalesRegister/UploadSalesLedgerFile.cshtml");
        }

        public IActionResult SalesLedgerUpload(string ticketId, string Edit)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            ViewBag.Message = ticketId;
            ViewBag.Edit = Edit;
            return View("~/Views/Admin/SalesRegister/UploadSalesLedgerFile.cshtml");
        }

        [HttpPost]
        public async Task<IActionResult> SalesLedgerUploadFile(string financialYear, string periodtype, string period, string requestNo, IFormFile gstFile)
        {
            //
            // Console.WriteLine("SalesLedgerUploadFileAsync");
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6"; // For active button styling
            ViewBag.requestNo = requestNo;
            ViewBag.Edit = string.IsNullOrEmpty(requestNo) ? "No" : "Yes";
            if (gstFile == null || gstFile.Length == 0)
            {
                ViewBag.ErrorMessage = "Please select a valid file.";
                return View("~/Views/Admin/SalesRegister/UploadSalesLedgerFile.cshtml");
            }
            string[] validMonths = null;            // Determine valid months based on periodType
            if (periodtype.Equals("Monthly", StringComparison.OrdinalIgnoreCase))
            {
                validMonths = new string[] { period }; // period in Nov-24

            }
            else // Quarterly  Q1-2025, Q2-2025, Q3-2025, Q4-2026  for FY 2025-26
            {
                var split = period.Split('-');

                string quarter = split[0]; // Q1,Q2,Q3,Q4
                string year = split[1];    // 2025,2025,2025,2026
                string shortYear = year.Substring(2); // "25","25", "25", "26"  


                switch (quarter.ToUpper())
                {
                    case "Q1":
                        validMonths = new string[] { $"Apr-{shortYear}", $"May-{shortYear}", $"Jun-{shortYear}" };
                        break;
                    case "Q2":
                        validMonths = new string[] { $"Jul-{shortYear}", $"Aug-{shortYear}", $"Sep-{shortYear}" };
                        break;
                    case "Q3":
                        validMonths = new string[] { $"Oct-{shortYear}", $"Nov-{shortYear}", $"Dec-{shortYear}" };
                        break;
                    case "Q4":
                        validMonths = new string[] { $"Jan-{shortYear}", $"Feb-{shortYear}", $"Mar-{shortYear}" };
                        break;
                    default:
                        throw new Exception($"Invalid quarter: '{quarter}'. Expected Q1 to Q4.");
                }
            }

            DataTable dataTable;
            string extension = Path.GetExtension(gstFile.FileName);
            string fileName = Path.GetFileName(gstFile.FileName);
            var _userGstin = MySession.Current.gstin; // Get the GSTIN from the session
                                                      //Console.WriteLine($"User GSTIN: {_userGstin}"); // Log the GSTIN for debugging
            if (extension == ".csv")
            {
                using (var stream = new MemoryStream())
                {
                    await gstFile.CopyToAsync(stream);
                    stream.Position = 0;
                    try
                    {
                        dataTable = ReadSalesLedgerCsvFile(stream, validMonths, _userGstin);
                        if (!ValidateSLColumnNames(dataTable))
                        {
                            ViewBag.ErrorMessage = "Please check columns names";
                            return View("~/Views/Admin/SalesRegister/UploadSalesLedgerFile.cshtml");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        return View("~/Views/Admin/SalesRegister/UploadSalesLedgerFile.cshtml");
                    }
                }
            }
            else if (extension == ".xlsx")
            {
                string sheetName = _configuration["SL_Invoice_Xlsx_SheetName"];
                using (var stream = new MemoryStream())
                {
                    await gstFile.CopyToAsync(stream);
                    stream.Position = 0;
                    try
                    {
                        dataTable = ReadSalesLedgerExcelFile(stream, validMonths, sheetName, _userGstin); //Change sheet name from hard-code to get from user  
                        if (!ValidateSLColumnNames(dataTable))
                        {
                            ViewBag.ErrorMessage = "Please check columns names";
                            return View("~/Views/Admin/SalesRegister/UploadSalesLedgerFile.cshtml");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        return View("~/Views/Admin/SalesRegister/UploadSalesLedgerFile.cshtml");
                    }
                }
            }
            else if (extension == ".xls")
            {
                string sheetName = _configuration["SL_Invoice_Xlsx_SheetName"];
                using (var stream = new MemoryStream())
                {
                    await gstFile.CopyToAsync(stream);
                    stream.Position = 0;
                    try
                    {
                        dataTable = ReadSalesLedgerXLSFile(stream, validMonths, sheetName, _userGstin); //Change sheet name from hard-code to get from user  
                        if (!ValidateSLColumnNames(dataTable))
                        {
                            ViewBag.ErrorMessage = "Please check columns names";
                            return View("~/Views/Admin/SalesRegister/UploadSalesLedgerFile.cshtml");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        return View("~/Views/Admin/SalesRegister/UploadSalesLedgerFile.cshtml");
                    }
                }
            }
            else
            {
                ViewBag.ErrorMessage = "Invalid file format. Please upload CSV/Xlsx/Xls file";
                return View("~/Views/Admin/SalesRegister/UploadSalesLedgerFile.cshtml");
            }

            //Console.WriteLine("Hello");
            string name = MySession.Current.UserName;
            string usergstin = MySession.Current.gstin;
            ViewBag.ticketId = string.IsNullOrEmpty(requestNo) ? GenerateSalesLedgerTicket() : requestNo;
            string ticketId = ViewBag.ticketId;
            try
            {
                //Console.WriteLine("Saving Sales Ledger Data");
                await _sLDataBusiness.SaveSalesLedgerDataAsync(dataTable, ticketId, usergstin);
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Error while saving Sales Ledger Data: " + ex.Message);
                ViewBag.ErrorMessage = $"{ex.Message} ";
                return View("~/Views/Admin/SalesRegister/UploadSalesLedgerFile.cshtml");
            }

            string Edit = ViewBag.Edit;
            DateTime? createdDate = DateTime.Now;
            if (ViewBag.Edit == "Yes")
            {
                var clientGSTIN = MySession.Current.gstin;
                //var ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(requestNo, clientGSTIN);
                var ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(requestNo, clientGSTIN);
                createdDate = ticket.RequestCreatedDate;
            }

            SalesLedgerTicketsModel data = new SalesLedgerTicketsModel();
            {
                data.RequestNumber = ticketId;
                data.ClientGstin = usergstin;
                data.ClientName = name;
                data.CLientEmail = MySession.Current.Email;
                data.RequestCreatedDate = createdDate;
                data.RequestUpdatedDate = DateTime.Now;
                data.FileName = fileName;
                data.Edit = Edit;
                data.FinancialYear = financialYear;
                data.PeriodType = periodtype;
                data.Period = period;
            }
            await _sLTicketsBusiness.SaveSLTicketAsync(data);
            return RedirectToAction("SalesLedgerUpload", "Admin", new { ticketId, Edit });
            // return View("~/Views/Admin/SalesRegister/UploadSalesLedgerFile.cshtml");
        }

        private string GenerateSalesLedgerTicket()
        {
            return "REQ_SL_" + DateTime.Now.ToString("yyyyMMddHHmmss");
        }
        private DataTable ReadSalesLedgerCsvFile(Stream stream, string[] period, string gstin)
        {
            //Console.WriteLine("Hi");
            var SLInvoice = _configuration["SLInvoice"];
            var SLInvoiceList = SLInvoice.Split(',').Select(x => x.Trim()).ToList();

            DataTable dt = new DataTable();
            int lineNumber = 0;

            int invNoColumnIndex = -1;
            int invoiceDateColumnIndex = -1;
            int userGstinIndex = -1;
            int supplierNameIndex = -1;
            int supplierGstinIndex = -1;
            int materialDescIndex = -1;

            List<int> columnMismatchRows = new List<int>();

            List<int> InvNoInvalideRows = new List<int>();
            List<int> invoiceDateInvalidRows = new List<int>();
            List<int> userGstinInvalidRows = new List<int>();
            List<int> supplierNameInvalidRows = new List<int>();
            List<int> supplierGstinInvalidRows = new List<int>();
            List<int> materialDescInvalidRows = new List<int>();

            string[] PeriodSub = period.Select(p => p.Substring(0, 3)).ToArray();
            string[] expectednumPeriod = period.Select(p => getNumber(p.Substring(0, 3))).ToArray();
            string[] periodYear = period.Select(p => p.Substring(p.Length - 2)).ToArray();
            // string UserGstin = gstin; // Remove all internal spaces and trim
            //Console.WriteLine("1");
            using (var reader = new StreamReader(stream))
            using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(reader))
            {
                var _SLinvoice = _configuration["SLInvoice"];
                var _SLinvoiceColumns = _SLinvoice.Split(',').Select(x => x.Trim()).ToList();

                parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = true;

                // Read header
                string[] headers = parser.ReadFields();
                lineNumber++;

                foreach (var header in headers)
                    dt.Columns.Add(header.Trim(), typeof(string));
                //S.NO	INV NO	DATE 	NAME OF CUSTOMERS	GST No	State to Supply(POS)	TAXABLE AMOUNT	 IGST	 SGST	CGST	TOTAL	GST Rates	HSN/SAC Code	Qty 	Units	Material Description	Type of Invoice

                // invno , date, nameof customer , gst no , material desc

                //Invoice number Index
                invNoColumnIndex = Array.FindIndex(headers, h => h.Trim().Equals(SLInvoiceList[1], StringComparison.OrdinalIgnoreCase));
                if (invNoColumnIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{SLInvoiceList[1]}' column.");
                //Invoice date Index
                invoiceDateColumnIndex = Array.FindIndex(headers, h => h.Trim().Equals(SLInvoiceList[2], StringComparison.OrdinalIgnoreCase));
                if (invoiceDateColumnIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{SLInvoiceList[2]}' column.");
                //User GSTIN Index
                userGstinIndex = Array.FindIndex(headers, h => h.Trim().Equals(SLInvoiceList[17], StringComparison.OrdinalIgnoreCase));
                if (userGstinIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{SLInvoiceList[17]}' column.");
                //Supplier Name Index
                supplierNameIndex = Array.FindIndex(headers, h => h.Trim().Equals(SLInvoiceList[3], StringComparison.OrdinalIgnoreCase));
                if (supplierNameIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{SLInvoiceList[3]}' column.");
                //Supplier GSTIN Index
                supplierGstinIndex = Array.FindIndex(headers, h => h.Trim().Equals(SLInvoiceList[4], StringComparison.OrdinalIgnoreCase));
                if (supplierGstinIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{SLInvoiceList[4]}' column.");
                //Material Description Index
                materialDescIndex = Array.FindIndex(headers, h => h.Trim().Equals(SLInvoiceList[15], StringComparison.OrdinalIgnoreCase));
                if (materialDescIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{SLInvoiceList[15]}' column.");
                //Console.WriteLine($"invNoColumnIndex - {invNoColumnIndex}");
                //Console.WriteLine($"invoiceDateColumnIndex - {invoiceDateColumnIndex}");
                //Console.WriteLine($"userGstinIndex - {userGstinIndex}");
                //Console.WriteLine($"supplierNameIndex - {supplierNameIndex}");
                //Console.WriteLine($"supplierGstinIndex - {supplierGstinIndex}");
                //Console.WriteLine($"materialDescIndex - {materialDescIndex}");
                //Console.ReadKey();



                //Console.WriteLine("2");
                // Read data rows
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    lineNumber++;
                    // Column count check
                    if (fields.Length != dt.Columns.Count)
                    {
                        columnMismatchRows.Add(lineNumber);
                        continue;
                    }

                    //Invoice number Validation
                    string invoiceNoValue = fields[invNoColumnIndex];
                    if (string.IsNullOrWhiteSpace(invoiceNoValue) || invoiceNoValue.Length > 25)
                    {
                        InvNoInvalideRows.Add(lineNumber);
                        continue;
                    }
                    //Invoice date Validation
                    string invoiceDateStr = fields[invoiceDateColumnIndex]?.Trim();
                    string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
                    if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
                    {
                        fields[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
                    }
                    else
                    {
                        //Console.WriteLine($"invoice date : {fields[invoiceDateColumnIndex]}");
                        invoiceDateInvalidRows.Add(lineNumber);
                        continue;
                    }
                    // user GSTIN Validation
                    string userGstinValue = fields[userGstinIndex]?.Trim();
                    //Console.WriteLine($"userGstinValue : {userGstinValue}");
                    if (string.IsNullOrWhiteSpace(userGstinValue) || userGstinValue.Length != 15 || userGstinValue != gstin)
                    {
                        userGstinInvalidRows.Add(lineNumber);
                        continue;
                    }
                    //Supplier Name Validation
                    string supplierNameValue = fields[supplierNameIndex];
                    if (supplierNameValue.Length > 100)
                    {
                        supplierNameInvalidRows.Add(lineNumber);
                        continue;
                    }
                    //Supplier GSTIN Validation
                    string supplierGstinValue = fields[supplierGstinIndex];
                    if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
                    {
                        supplierGstinInvalidRows.Add(lineNumber);
                        continue;
                    }
                    //Material Description Validation
                    string materialDescValue = fields[materialDescIndex];
                    if (materialDescValue.Length > 250)
                    {
                        materialDescInvalidRows.Add(lineNumber);
                        continue;
                    }

                    dt.Rows.Add(fields);
                    //Console.WriteLine("3");
                }
            }

            // Handle mismatch errors
            if (columnMismatchRows.Count > 0 || InvNoInvalideRows.Count > 0 || userGstinInvalidRows.Count > 0 || invoiceDateInvalidRows.Count > 0 || supplierNameInvalidRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || materialDescInvalidRows.Count > 0)
            {
                var errorMsg = "";

                if (columnMismatchRows.Count > 0)
                    errorMsg += $"Column count mismatch at line(s): {string.Join(", ", columnMismatchRows)}. Expected column count: {dt.Columns.Count}.";
                if (InvNoInvalideRows.Count > 0)
                    errorMsg += $"\n Invalid Invoice Number at line(s): {string.Join(", ", InvNoInvalideRows)}.";
                if (userGstinInvalidRows.Count > 0)
                    errorMsg += $"\n Invalid User GSTIN at line(s): {string.Join(", ", userGstinInvalidRows)}.";
                if (invoiceDateInvalidRows.Count > 0)
                    errorMsg += $"\n Invoice date format mismatch at line(s): {string.Join(", ", invoiceDateInvalidRows)}. Expected Format : dd-MM-yyyy hh:mm:ss tt .";
                if (supplierNameInvalidRows.Count > 0)
                    errorMsg += $"\n The Supplier Name field allows a maximum of 100 characters.Invalid Supplier Name at line(s) : {string.Join(", ", supplierNameInvalidRows)}.";
                if (supplierGstinInvalidRows.Count > 0)
                    errorMsg += $"\n Invalid Supplier GSTIN at line(s): {string.Join(", ", supplierGstinInvalidRows)}.";
                if (materialDescInvalidRows.Count > 0)
                    errorMsg += $"\n The Material Description field allows a maximum of 250 characters.Invalid Material Description at line(s) : {string.Join(", ", materialDescInvalidRows)}.";
                //Console.WriteLine("4");
                throw new Exception("Failed to insert data: \n" + errorMsg);
            }

            return dt;
        }
        private DataTable ReadSalesLedgerExcelFile(Stream stream, string[] period, string sheetName, string gstin)
        {
            var SLInvoice = _configuration["SLInvoice"];
            var SLInvoiceList = SLInvoice.Split(',').Select(x => x.Trim()).ToList();

            DataTable dt = new DataTable();
            int lineNumber = 0;

            int userGstinIndex = -1;
            int invNoColumnIndex = -1;
            int invoiceDateColumnIndex = -1;
            int supplierNameIndex = -1;
            int supplierGstinIndex = -1;
            int materialDescIndex = -1;

            List<int> columnMismatchRows = new List<int>();

            List<int> userGstinInvalidRows = new List<int>();
            List<int> InvNoInvalideRows = new List<int>();
            List<int> invoiceDateInvalidRows = new List<int>();
            List<int> supplierNameInvalidRows = new List<int>();
            List<int> supplierGstinInvalidRows = new List<int>();
            List<int> materialDescInvalidRows = new List<int>();

            string[] PeriodSub = period.Select(p => p.Substring(0, 3)).ToArray();
            string[] expectednumPeriod = period.Select(p => getNumber(p.Substring(0, 3))).ToArray();
            string[] periodYear = period.Select(p => p.Substring(p.Length - 2)).ToArray();
            using (var workbook = new XLWorkbook(stream))
            {
                //var worksheet = workbook.Worksheets.First();
                var worksheet = workbook.Worksheet(sheetName);
                if (worksheet == null)
                    throw new Exception($"Sheet '{sheetName}' not found.");
                var rows = worksheet.RowsUsed().ToList();

                if (rows.Count == 0)
                    throw new Exception("Excel file is empty.");

                // Read header
                var headerRow = rows[0];
                var headers = headerRow.Cells().Select(c => c.GetString().Trim()).ToArray();

                foreach (var header in headers)
                    dt.Columns.Add(header, typeof(string));

                //User GSTIN Index
                userGstinIndex = Array.FindIndex(headers, h => h.Trim().Equals(SLInvoiceList[17], StringComparison.OrdinalIgnoreCase));
                if (userGstinIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{SLInvoiceList[17]}' column.");
                //Invoice number Index
                invNoColumnIndex = Array.FindIndex(headers, h => h.Trim().Equals(SLInvoiceList[1], StringComparison.OrdinalIgnoreCase));
                if (invNoColumnIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{SLInvoiceList[1]}' column.");
                //Invoice date Index
                invoiceDateColumnIndex = Array.FindIndex(headers, h => h.Trim().Equals(SLInvoiceList[2], StringComparison.OrdinalIgnoreCase));
                if (invoiceDateColumnIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{SLInvoiceList[2]}' column.");
                //Supplier Name Index
                supplierNameIndex = Array.FindIndex(headers, h => h.Trim().Equals(SLInvoiceList[3], StringComparison.OrdinalIgnoreCase));
                if (supplierNameIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{SLInvoiceList[3]}' column.");
                //Supplier GSTIN Index
                supplierGstinIndex = Array.FindIndex(headers, h => h.Trim().Equals(SLInvoiceList[4], StringComparison.OrdinalIgnoreCase));
                if (supplierGstinIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{SLInvoiceList[4]}' column.");
                //Material Description Index
                materialDescIndex = Array.FindIndex(headers, h => h.Trim().Equals(SLInvoiceList[15], StringComparison.OrdinalIgnoreCase));
                if (materialDescIndex == -1)
                    throw new Exception($"The CSV file does not contain an '{SLInvoiceList[15]}' column.");


                // Read data rows
                for (int i = 1; i < rows.Count; i++)
                {
                    int excelLineNumber = i + 1; // Line number as seen in Excel (header is line 1)
                    var row = rows[i];
                    var cells = row.Cells().Select(c => c.GetString().Trim()).ToArray();
                    // Column count check
                    if (cells.Length != dt.Columns.Count)
                    {
                        columnMismatchRows.Add(excelLineNumber);
                        continue;
                    }
                    // Invoice number Validation
                    string invoiceNoValue = cells[invNoColumnIndex];
                    if (string.IsNullOrWhiteSpace(invoiceNoValue) || invoiceNoValue.Length > 25)
                    {
                        InvNoInvalideRows.Add(excelLineNumber);
                        continue;
                    }
                    // Invoice date Validation
                    string invoiceDateStr = cells[invoiceDateColumnIndex]?.Trim();
                    string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
                    if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
                    {
                        cells[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
                    }
                    else
                    {
                        //Console.WriteLine($"invoice date : {fields[invoiceDateColumnIndex]}");
                        invoiceDateInvalidRows.Add(excelLineNumber);
                        continue;
                    }
                    // user GSTIN Validation
                    string userGstinValue = cells[userGstinIndex]?.Trim();
                    if (string.IsNullOrWhiteSpace(userGstinValue) || userGstinValue.Length != 15 || userGstinValue != gstin)
                    {
                        userGstinInvalidRows.Add(excelLineNumber);
                        continue;
                    }
                    // Supplier Name Validation
                    string supplierNameValue = cells[supplierNameIndex];
                    if (supplierNameValue.Length > 100)
                    {
                        supplierNameInvalidRows.Add(excelLineNumber);
                        continue;
                    }
                    // Supplier GSTIN Validation
                    string supplierGstinValue = cells[supplierGstinIndex];
                    if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
                    {
                        supplierGstinInvalidRows.Add(excelLineNumber);
                        continue;
                    }
                    // Material Description Validation
                    string materialDescValue = cells[materialDescIndex];
                    if (materialDescValue.Length > 250)
                    {
                        materialDescInvalidRows.Add(excelLineNumber);
                        continue;
                    }

                    dt.Rows.Add(cells);
                }

            }
            // Handle mismatch errors
            if (columnMismatchRows.Count > 0 || userGstinInvalidRows.Count > 0 || InvNoInvalideRows.Count > 0 || invoiceDateInvalidRows.Count > 0 || supplierNameInvalidRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || materialDescInvalidRows.Count > 0)
            {
                var errorMsg = "";

                if (userGstinInvalidRows.Count > 0)
                    errorMsg += $"\n Invalid User GSTIN at line(s): {string.Join(", ", userGstinInvalidRows)}.";
                if (columnMismatchRows.Count > 0)
                    errorMsg += $"Column count mismatch at line(s): {string.Join(", ", columnMismatchRows)}. Expected column count: {dt.Columns.Count}.";
                if (InvNoInvalideRows.Count > 0)
                    errorMsg += $"\n Invalid Invoice Number at line(s): {string.Join(", ", InvNoInvalideRows)}.";
                if (invoiceDateInvalidRows.Count > 0)
                    errorMsg += $"\n Invoice date format mismatch at line(s): {string.Join(", ", invoiceDateInvalidRows)}. Expected Format : dd-MM-yyyy hh:mm:ss tt .";
                if (supplierNameInvalidRows.Count > 0)
                    errorMsg += $"\n The Supplier Name field allows a maximum of 100 characters.Invalid Supplier Name at line(s) : {string.Join(", ", supplierNameInvalidRows)}.";
                if (supplierGstinInvalidRows.Count > 0)
                    errorMsg += $"\n Invalid Supplier GSTIN at line(s): {string.Join(", ", supplierGstinInvalidRows)}.";
                if (materialDescInvalidRows.Count > 0)
                    errorMsg += $"\n The Material Description field allows a maximum of 250 characters.Invalid Material Description at line(s) : {string.Join(", ", materialDescInvalidRows)}.";

                throw new Exception("Failed to insert data: \n" + errorMsg);
            }

            return dt;
        }
        private DataTable ReadSalesLedgerXLSFile(Stream stream, string[] period, string sheetName, string gstin)
        {
            var SLInvoice = _configuration["SLInvoice"];
            var SLInvoiceList = SLInvoice.Split(',').Select(x => x.Trim()).ToList();

            DataTable dt = new DataTable();
            int lineNumber = 0;

            int userGstinIndex = -1;
            int invNoColumnIndex = -1;
            int invoiceDateColumnIndex = -1;
            int supplierNameIndex = -1;
            int supplierGstinIndex = -1;
            int materialDescIndex = -1;

            List<int> columnMismatchRows = new List<int>();

            List<int> userGstinInvalidRows = new List<int>();
            List<int> InvNoInvalideRows = new List<int>();
            List<int> invoiceDateInvalidRows = new List<int>();
            List<int> supplierNameInvalidRows = new List<int>();
            List<int> supplierGstinInvalidRows = new List<int>();
            List<int> materialDescInvalidRows = new List<int>();

            string[] PeriodSub = period.Select(p => p.Substring(0, 3)).ToArray();
            string[] expectednumPeriod = period.Select(p => getNumber(p.Substring(0, 3))).ToArray();
            string[] periodYear = period.Select(p => p.Substring(p.Length - 2)).ToArray();

            HSSFWorkbook workbook = new HSSFWorkbook(stream);
            ISheet sheet = workbook.GetSheet(sheetName) ?? throw new Exception($"Sheet '{sheetName}' not found.");

            if (sheet.PhysicalNumberOfRows == 0)
                throw new Exception("Excel file is empty.");

            // Read header
            IRow headerRow = sheet.GetRow(0);
            int columnCount = headerRow.LastCellNum;
            string[] headers = new string[columnCount];

            for (int i = 0; i < columnCount; i++)
            {
                headers[i] = headerRow.GetCell(i)?.ToString().Trim() ?? $"Column{i}";
                dt.Columns.Add(headers[i], typeof(string));
            }

            // Index mapping and validation
            invNoColumnIndex = FindColumnIndex(headers, SLInvoiceList[1]);
            invoiceDateColumnIndex = FindColumnIndex(headers, SLInvoiceList[2]);
            userGstinIndex = FindColumnIndex(headers, SLInvoiceList[17]);
            supplierNameIndex = FindColumnIndex(headers, SLInvoiceList[3]);
            supplierGstinIndex = FindColumnIndex(headers, SLInvoiceList[4]);
            materialDescIndex = FindColumnIndex(headers, SLInvoiceList[15]);

            // Process rows
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;

                string[] cells = new string[columnCount];
                for (int j = 0; j < columnCount; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell == null)
                    {
                        cells[j] = string.Empty;
                        continue;
                    }

                    if (j == invoiceDateColumnIndex && cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
                    {
                        DateTime DT = (DateTime)cell.DateCellValue;
                        cells[j] = DT.ToString("dd-MM-yyyy hh:mm:ss tt"); // Match your allowed format
                    }
                    else
                    {
                        cells[j] = cell.ToString().Trim();
                    }
                }


                int excelLineNumber = i + 1;

                // Column count check
                if (cells.Length != dt.Columns.Count)
                {
                    columnMismatchRows.Add(excelLineNumber);
                    continue;
                }
                // Invoice number Validation
                string invoiceNoValue = cells[invNoColumnIndex];
                if (string.IsNullOrWhiteSpace(invoiceNoValue) || invoiceNoValue.Length > 25)
                {
                    InvNoInvalideRows.Add(excelLineNumber);
                    continue;
                }
                // Invoice date Validation
                string invoiceDateStr = cells[invoiceDateColumnIndex]?.Trim();
                string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
                if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
                {
                    cells[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
                }
                else
                {
                    //Console.WriteLine($"invoice date : {fields[invoiceDateColumnIndex]}");
                    invoiceDateInvalidRows.Add(excelLineNumber);
                    continue;
                }
                // user GSTIN Validation
                string userGstinValue = cells[userGstinIndex]?.Trim();
                if (string.IsNullOrWhiteSpace(userGstinValue) || userGstinValue.Length != 15 || userGstinValue != gstin)
                {
                    userGstinInvalidRows.Add(excelLineNumber);
                    continue;
                }
                // Supplier Name Validation
                string supplierNameValue = cells[supplierNameIndex];
                if (supplierNameValue.Length > 100)
                {
                    supplierNameInvalidRows.Add(excelLineNumber);
                    continue;
                }
                // Supplier GSTIN Validation
                string supplierGstinValue = cells[supplierGstinIndex];
                if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
                {
                    supplierGstinInvalidRows.Add(excelLineNumber);
                    continue;
                }
                // Material Description Validation
                string materialDescValue = cells[materialDescIndex];
                if (materialDescValue.Length > 250)
                {
                    materialDescInvalidRows.Add(excelLineNumber);
                    continue;
                }

                dt.Rows.Add(cells);


            }
            // Handle mismatch errors
            if (columnMismatchRows.Count > 0 || userGstinInvalidRows.Count > 0 || InvNoInvalideRows.Count > 0 || invoiceDateInvalidRows.Count > 0 || supplierNameInvalidRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || materialDescInvalidRows.Count > 0)
            {
                var errorMsg = "";

                if (userGstinInvalidRows.Count > 0)
                    errorMsg += $"\n Invalid User GSTIN at line(s): {string.Join(", ", userGstinInvalidRows)}.";
                if (columnMismatchRows.Count > 0)
                    errorMsg += $"Column count mismatch at line(s): {string.Join(", ", columnMismatchRows)}. Expected column count: {dt.Columns.Count}.";
                if (InvNoInvalideRows.Count > 0)
                    errorMsg += $"\n Invalid Invoice Number at line(s): {string.Join(", ", InvNoInvalideRows)}.";
                if (invoiceDateInvalidRows.Count > 0)
                    errorMsg += $"\n Invoice date formate mismatch at line(s): {string.Join(", ", invoiceDateInvalidRows)}. Expected Formate : dd-MM-yyyy hh:mm:ss tt .";
                if (supplierNameInvalidRows.Count > 0)
                    errorMsg += $"\n The Supplier Name field allows a maximum of 100 characters.Invalid Supplier Name at line(s) : {string.Join(", ", supplierNameInvalidRows)}.";
                if (supplierGstinInvalidRows.Count > 0)
                    errorMsg += $"\n Invalid Supplier GSTIN at line(s): {string.Join(", ", supplierGstinInvalidRows)}.";
                if (materialDescInvalidRows.Count > 0)
                    errorMsg += $"\n The Material Description field allows a maximum of 250 characters.Invalid Material Description at line(s) : {string.Join(", ", materialDescInvalidRows)}.";

                throw new Exception("Failed to insert data: \n" + errorMsg);
            }

            return dt;
        }
        private bool ValidateSLColumnNames(DataTable dataTable)
        {
            var _SLinvoice = _configuration["SLInvoice"];
            var _SLinvoiceColumns = _SLinvoice.Split(',').Select(x => x.Trim()).ToList();
            //foreach (var list in _SLinvoiceColumns)
            //{
            //    Console.WriteLine(list);
            //}

            var uploadedColumns = dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName.ToLower()).ToList();
            //foreach (var col in _SLinvoiceColumns)
            //{
            //    Console.WriteLine(col.ToLower());
            //}
            return _SLinvoiceColumns.All(col => uploadedColumns.Contains(col.ToLower()));
        }

        #endregion

        #region SalesLedgerCurrentRequestsCSV
        public async Task<IActionResult> SalesLedgerCurrentRequestsCSV(DateTime? fromdate, DateTime? todate)
		{
			DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
			DateTime toDateTime = todate ?? DateTime.Now; // Default to today

			//await TjCaptions("OpenTasks");
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act6";
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                return View("~/Views/Admin/SalesLedgerCurrentRequests/SalesLedgerCurrentRequestsCSV.cshtml");
            }

   //         var email = MySession.Current.Email; // Get the email from session
			//var clients = await _userBusiness.GetAdminClients(email);
			//string[] gstinArray = clients.Select(c => c.ClientGSTIN).ToArray();
			//var Tickets = await _sLTicketsBusiness.GetClientsOpenTicketsStatusAsync(gstinArray, fromDateTime, toDateTime);

            var Tickets = await _sLTicketsBusiness.GetOpenTicketsStatusAsync(MySession.Current.gstin, fromDateTime, toDateTime); //(ViewBag.UserGSTIN); // Fetch pending tickets based on user_gstin

            ViewBag.Tickets = Tickets;

			return View("~/Views/Admin/SalesLedgerCurrentRequests/SalesLedgerCurrentRequestsCSV.cshtml");
		}

		public async Task<IActionResult> CompareSalesLedgerGSTCSV(string ticketnumber, string ClientGSTIN, string fromDate, string toDate)
		{
			// _logger.LogInformation($"Ticket Number: {ticketnumber}");
			//based on this ticket number fetch data from Purchase Ticket table and store
			var Ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);

			ViewBag.Ticket = Ticket;
			ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            ViewBag.fromDate = fromDate;
			ViewBag.toDate = toDate;

			return View("~/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
		}

		[HttpPost]
		public async Task<IActionResult> CompareSalesLedgerCSV(IFormFile EInvoice, IFormFile EWayBill, string ticketnumber, string ClientGSTIN)
		{
			//await TjCaptions("CompareSalesLedgerCSV");
			ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            var Ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
			ViewBag.Ticket = Ticket;
			var matchTypes = _configuration["SLMatchTypes"];
			var matchTypeList = matchTypes.Split(',').Select(x => x.Trim()).ToList();
			ViewBag.MatchType = matchTypeList;
			// Console.WriteLine($"EInvoice_AdminFileName = {Path.GetFileName(EInvoice.FileName)}");    
			if (EInvoice == null || EInvoice.Length == 0 || EWayBill == null || EWayBill.Length == 0)
			{
				ViewBag.ErrorMessage = "Please select a valid file.";
				return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
			}

			DataTable EInvoice_dataTable = null;
			string EInvoice_extension = Path.GetExtension(EInvoice.FileName);
			string EInvoice_AdminFileName = Path.GetFileName(EInvoice.FileName);
			//Console.WriteLine($"AdminFileName: {AdminFileName}"); // Log the file name for debugging
			DataTable EWayBill_dataTable = null;
			string EWayBill_extension = Path.GetExtension(EWayBill.FileName);
			string EWayBill_AdminFileName = Path.GetFileName(EWayBill.FileName);

			bool isEInvoiceInvalid = EInvoice_extension != ".csv" && EInvoice_extension != ".xls" && EInvoice_extension != ".xlsx";
			bool isEWayBillInvalid = EWayBill_extension != ".csv" && EWayBill_extension != ".xls" && EWayBill_extension != ".xlsx";

			if (isEInvoiceInvalid && isEWayBillInvalid)
			{
				// ❌ Both files are invalid — show error
				ViewBag.ErrorMessage = "Both Files format are InValid . Please upload CSV / Xlsx / Xls files";
				return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
			}
			else if (isEInvoiceInvalid)
			{
				// ❌ Only E-Invoice is invalid
				ViewBag.ErrorMessage = "E-Invoice file format is invalid. Please upload CSV / Xlsx / Xls files";
				return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
			}
			else if (isEWayBillInvalid)
			{
				// ❌ Only E-Way Bill is invalid
				ViewBag.ErrorMessage = "E-Way Bill file format is invalid. Please upload CSV / Xlsx / Xls files";
				return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
			}
			else
			{
				if (EInvoice_extension == ".csv")
				{
					using (var stream = new MemoryStream())
					{
						await EInvoice.CopyToAsync(stream);
						stream.Position = 0;
						try
						{
							EInvoice_dataTable = ReadEInvoiceCsvFile(stream, ClientGSTIN);
							if (!ValidateEInvoiceColumnNames(EInvoice_dataTable))
							{
								ViewBag.ErrorMessage = "Please check EInvoice File columns names";
								return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
							}
						}
						catch (Exception ex)
						{
							ViewBag.ErrorMessage = ex.Message;
							return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
						}
					}
				}
				if (EInvoice_extension == ".xlsx")
				{
					string sheetName = _configuration["SL_EInvoice_Xlsx_SheetName"];
					using (var stream = new MemoryStream())
					{
						await EInvoice.CopyToAsync(stream);
						stream.Position = 0;
						try
						{
							EInvoice_dataTable = ReadEInvoiceExcelFile(stream, sheetName, ClientGSTIN);
							if (!ValidateEInvoiceColumnNames(EInvoice_dataTable))
							{
								ViewBag.ErrorMessage = "Please check EInvoice File columns names";
								return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
							}
						}
						catch (Exception ex)
						{
							ViewBag.ErrorMessage = ex.Message;
							return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
						}
					}
				}
				if (EInvoice_extension == ".xls")
				{
					string sheetName = _configuration["SL_EInvoice_Xlsx_SheetName"];
					using (var stream = new MemoryStream())
					{
						await EInvoice.CopyToAsync(stream);
						stream.Position = 0;
						try
						{
							EInvoice_dataTable = ReadEInvoiceXLSFile(stream, sheetName, ClientGSTIN);
							if (!ValidateEInvoiceColumnNames(EInvoice_dataTable))
							{
								ViewBag.ErrorMessage = "Please check EInvoice File columns names";
								return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
							}
						}
						catch (Exception ex)
						{
							ViewBag.ErrorMessage = ex.Message;
							return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
						}
					}
				}

				if (EWayBill_extension == ".csv")
				{
					using (var stream = new MemoryStream())
					{
						await EWayBill.CopyToAsync(stream);
						stream.Position = 0;
						try
						{
							EWayBill_dataTable = ReadEWayBillCsvFile(stream, ClientGSTIN);
							if (!ValidateEWayBillColumnNames(EWayBill_dataTable))
							{
								ViewBag.ErrorMessage = "Please check EWayBill File columns names";
								return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
							}
						}
						catch (Exception ex)
						{
							ViewBag.ErrorMessage = ex.Message;
							return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
						}
					}
				}
				if (EWayBill_extension == ".xlsx")
				{
					string sheetName = _configuration["SL_EWayBill_Xlsx_SheetName"];
					using (var stream = new MemoryStream())
					{
						await EWayBill.CopyToAsync(stream);
						stream.Position = 0;
						try
						{
							EWayBill_dataTable = ReadEwayBillExcelFile(stream, sheetName, ClientGSTIN);
							if (!ValidateEWayBillColumnNames(EWayBill_dataTable))
							{
								ViewBag.ErrorMessage = "Please check EWayBill File columns names";
								return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
							}
						}
						catch (Exception ex)
						{
							ViewBag.ErrorMessage = ex.Message;
							return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
						}
					}
				}
				if (EWayBill_extension == ".xls")
				{
					string sheetName = _configuration["SL_EWayBill_Xlsx_SheetName"];
					using (var stream = new MemoryStream())
					{
						await EWayBill.CopyToAsync(stream);
						stream.Position = 0;
						try
						{
							EWayBill_dataTable = ReadEwayBillXLSFile(stream, sheetName, ClientGSTIN);
							if (!ValidateEWayBillColumnNames(EWayBill_dataTable))
							{
								ViewBag.ErrorMessage = "Please check EWayBill File columns names";
								return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
							}
						}
						catch (Exception ex)
						{
							ViewBag.ErrorMessage = ex.Message;
							return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
						}
					}
				}
			}

			try
			{
				//await _sLDataBusiness.SaveSalesLedgerDataAsync(dataTable, ticketId, usergstin);
				await _sLEInvoiceBusiness.SaveEInvoiceDataAsync(EInvoice_dataTable, ticketnumber, ClientGSTIN);
				await _sLEWayBillBusiness.SaveEWayBillDataAsync(EWayBill_dataTable, ticketnumber, ClientGSTIN);
				
			}
			catch (Exception ex)
			{
				ViewBag.ErrorMessage = $"{ex.Message} ";
				return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTCSV.cshtml");
			}

			// Comparison logic (unchanged)
			DataTable SLInvoice = await _sLDataBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
			DataTable SLWayBill = await _sLEWayBillBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
			DataTable SLEInvoice = await _sLEInvoiceBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
			//Console.WriteLine("SLInvoice Columns: " + string.Join(", ", SLInvoice.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
			//Console.WriteLine("SLWayBill Columns: " + string.Join(", ", SLWayBill.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
			//Console.WriteLine("SLEInvoice Columns: " + string.Join(", ", SLEInvoice.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));




			//await _sLComparedDataBusiness.CompareDataAsync(SLInvoice, SLWayBill, SLEInvoice);
			await _sLComparedDataBusiness.CompareDataAsync(SLInvoice, SLWayBill, SLEInvoice);
			await _sLTicketsBusiness.UpdateSLTicketAsync(ticketnumber, ClientGSTIN, EInvoice_AdminFileName, EWayBill_AdminFileName);


			var data = await _sLComparedDataBusiness.GetComparedDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);

			
			ViewBag.ReportDataList = data;
			//_logger.LogInformation("Total Compared Data Rows: " + data.Count);

			// store number and sum in a model and save it in 
			ViewBag.Summary = GenerateSLSummary(data);
			ViewBag.GrandTotal = getGrandTotal(data);
			//var summary = ViewBag.Summary;
		   // Console.WriteLine($"Summary : {summary.catagory1InvoiceSum}");
			return View("~/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedger.cshtml");
		}

		private DataTable ReadEInvoiceCsvFile(Stream stream, string gstin)
		{

			var SLEInvoiceColumns = _configuration["SLEInvoice"].Split(',').Select(x => x.Trim()).ToList();

			DataTable dt = new DataTable();
			int lineNumber = 0;
			//suppliergstin,supplier name,invoiceno, invoice date,E-Commerce GSTIN,IRN date ,GSTR-1 auto-population/ deletion upon cancellation date	

			List<int> columnMismatchRows = new List<int>();
			List<int> supplierGstinInvalidRows = new List<int>();
			List<int> supplierNameInvalidRows = new List<int>();
			List<int> invoiceNumberInvalidRows = new List<int>();
			List<int> invoiceDateInvalidRows = new List<int>();
			List<int> eCommerceGstinInvalidRows = new List<int>();
			List<int> inrDateInvalidRows = new List<int>();
			List<int> gstr1AutoPopulationInvalidRows = new List<int>();

			using (var reader = new StreamReader(stream))
			using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(reader))
			{
				parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
				parser.SetDelimiters(",");
				parser.HasFieldsEnclosedInQuotes = true;

				// Read headers
				string[] headers = parser.ReadFields();
				lineNumber++;
				foreach (string header in headers)
				{
					dt.Columns.Add(header.Trim(), typeof(string));
				}
				// Index mapping and validation
				int supplierGstinIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[0]);
				int supplierNameColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[1]);
				int invoiceNumberColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[2]);
				int invoiceDateColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[3]);
				int eCommerceGstinColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[9]);
				int inrDateColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[17]);
				int gstr1AutoPopulationColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[19]);

				// Read data rows
				while (!parser.EndOfData)
				{
					string[] fields = parser.ReadFields();
					lineNumber++;

					if (fields.Length != dt.Columns.Count)
					{
						columnMismatchRows.Add(lineNumber);
						continue;
					}
					// Supplier GSTIN Validation
					string supplierGstinValue = fields[supplierGstinIndex];
					if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
					{
						supplierGstinInvalidRows.Add(lineNumber);
						continue;
					}
					// Supplier Name Validation
					string supplierNameValue = fields[supplierNameColumnIndex]?.Trim();
					if (supplierNameValue.Length > 100)
					{
						supplierNameInvalidRows.Add(lineNumber);
						continue;
					}
					// Invoice Number Validation
					string invoiceNumberValue = fields[invoiceNumberColumnIndex]?.Trim();
					if (string.IsNullOrWhiteSpace(invoiceNumberValue) || invoiceNumberValue.Length > 25)
					{
						invoiceNumberInvalidRows.Add(lineNumber);
						continue;
					}
					// Invoice Date Validation
					string invoiceDateStr = fields[invoiceDateColumnIndex]?.Trim();
					string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
					if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
					{
						fields[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						invoiceDateInvalidRows.Add(lineNumber);
						continue;
					}
					// E-Commerce GSTIN Validation
					string eCommerceGstinValue = fields[eCommerceGstinColumnIndex]?.Trim();
					if (string.IsNullOrWhiteSpace(eCommerceGstinValue) || eCommerceGstinValue.Length != 15)
					{
						eCommerceGstinInvalidRows.Add(lineNumber);
						continue;
					}
					// IRN Date Validation
					string inrDateStr = fields[inrDateColumnIndex]?.Trim();
					if (DateTime.TryParseExact(inrDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime inrDate))
					{
						fields[inrDateColumnIndex] = inrDate.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						inrDateInvalidRows.Add(lineNumber);
						continue;
					}
					// GSTR-1 auto-population/ deletion upon cancellation date Validation
					string gstr1AutoPopulationValue = fields[gstr1AutoPopulationColumnIndex]?.Trim();
					if (DateTime.TryParseExact(inrDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime gstr1Date))
					{
						fields[gstr1AutoPopulationColumnIndex] = gstr1Date.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						gstr1AutoPopulationInvalidRows.Add(lineNumber);
						continue;
					}


					dt.Rows.Add(fields);
				}
			}
			// Handle mismatch errors
			if (columnMismatchRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || supplierNameInvalidRows.Count > 0 || invoiceNumberInvalidRows.Count > 0 || invoiceDateInvalidRows.Count > 0 || eCommerceGstinInvalidRows.Count > 0 || inrDateInvalidRows.Count > 0 || gstr1AutoPopulationInvalidRows.Count > 0)
			{
				string error = "";
				if (columnMismatchRows.Count > 0)
					error += $"Column mismatch at rows: {string.Join(", ", columnMismatchRows)}\n";
				if (supplierGstinInvalidRows.Count > 0)
					error += $"Invalid Supplier GSTIN at rows: {string.Join(", ", supplierGstinInvalidRows)}\n";
				if (supplierNameInvalidRows.Count > 0)
					error += $"\n The Supplier Name field allows a maximum of 100 characters.Invalid Supplier Name at line(s) : {string.Join(", ", supplierNameInvalidRows)}.";
				if (invoiceNumberInvalidRows.Count > 0)
					error += $"\n The Invoice Number field allows a maximum of 25 characters.Invalid Invoice Number at line(s) : {string.Join(", ", invoiceNumberInvalidRows)}.";
				if (invoiceDateInvalidRows.Count > 0)
					error += $"\n Invalid Invoice Date at line(s) : {string.Join(", ", invoiceDateInvalidRows)}.";
				if (eCommerceGstinInvalidRows.Count > 0)
					error += $"\n Invalid E-Commerce GSTIN at line(s) : {string.Join(", ", eCommerceGstinInvalidRows)}.";
				if (inrDateInvalidRows.Count > 0)
					error += $"\n Invalid IRN Date at line(s) : {string.Join(", ", inrDateInvalidRows)}.";
				if (gstr1AutoPopulationInvalidRows.Count > 0)
					error += $"\n Invalid GSTR-1 auto-population/ deletion upon cancellation date at line(s) : {string.Join(", ", gstr1AutoPopulationInvalidRows)}.";

				throw new Exception("Failed to insert data: \n" + error);

			}
			return dt;
		}
		private DataTable ReadEInvoiceExcelFile(Stream stream, string sheetName, string gstin)
		{
			var SLEInvoiceColumns = _configuration["SLEInvoice"].Split(',').Select(x => x.Trim()).ToList();

			DataTable dt = new DataTable();
			int lineNumber = 0;
			//suppliergstin,supplier name,invoiceno, invoice date,E-Commerce GSTIN,IRN date ,GSTR-1 auto-population/ deletion upon cancellation date	

			List<int> columnMismatchRows = new List<int>();
			List<int> supplierGstinInvalidRows = new List<int>();
			List<int> supplierNameInvalidRows = new List<int>();
			List<int> invoiceNumberInvalidRows = new List<int>();
			List<int> invoiceDateInvalidRows = new List<int>();
			List<int> eCommerceGstinInvalidRows = new List<int>();
			List<int> inrDateInvalidRows = new List<int>();
			List<int> gstr1AutoPopulationInvalidRows = new List<int>();

			using (var workbook = new XLWorkbook(stream))
			{
				//var worksheet = workbook.Worksheets.First();
				var worksheet = workbook.Worksheet(sheetName);
				if (worksheet == null)
					throw new Exception($"Sheet '{sheetName}' not found.");
				var rows = worksheet.RowsUsed().ToList();

				if (rows.Count == 0)
					throw new Exception("Excel file is empty.");

				// Read header
				var headerRow = rows[0];
				var headers = headerRow.Cells().Select(c => c.GetString().Trim()).ToArray();

				foreach (var header in headers)
					dt.Columns.Add(header, typeof(string));

				// Index mapping and validation
				int supplierGstinIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[0]);
				int supplierNameColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[1]);
				int invoiceNumberColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[2]);
				int invoiceDateColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[3]);
				int eCommerceGstinColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[9]);
				int inrDateColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[17]);
				int gstr1AutoPopulationColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[19]);

				// Read data rows
				for (int i = 1; i < rows.Count; i++)
				{
					int excelLineNumber = i + 1; // Line number as seen in Excel (header is line 1)
					var row = rows[i];
					var cells = row.Cells().Select(c => c.GetString().Trim()).ToArray();
					// Column count check
					if (cells.Length != dt.Columns.Count)
					{
						columnMismatchRows.Add(excelLineNumber);
						continue;
					}
					// Supplier GSTIN Validation
					string supplierGstinValue = cells[supplierGstinIndex];
					if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
					{
						supplierGstinInvalidRows.Add(excelLineNumber);
						continue;
					}
					// Supplier Name Validation
					string supplierNameValue = cells[supplierNameColumnIndex]?.Trim();
					if (supplierNameValue.Length > 100)
					{
						supplierNameInvalidRows.Add(excelLineNumber);
						continue;
					}
					// Invoice Number Validation
					string invoiceNumberValue = cells[invoiceNumberColumnIndex]?.Trim();
					if (string.IsNullOrWhiteSpace(invoiceNumberValue) || invoiceNumberValue.Length > 25)
					{
						invoiceNumberInvalidRows.Add(excelLineNumber);
						continue;
					}
					// Invoice Date Validation
					string invoiceDateStr = cells[invoiceDateColumnIndex]?.Trim();
					string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
					if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
					{
						cells[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						invoiceDateInvalidRows.Add(excelLineNumber);
						continue;
					}
					// E-Commerce GSTIN Validation
					string eCommerceGstinValue = cells[eCommerceGstinColumnIndex]?.Trim();
					if (string.IsNullOrWhiteSpace(eCommerceGstinValue) || eCommerceGstinValue.Length != 15)
					{
						eCommerceGstinInvalidRows.Add(excelLineNumber);
						continue;
					}
					// IRN Date Validation
					string inrDateStr = cells[inrDateColumnIndex]?.Trim();
					if (DateTime.TryParseExact(inrDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime inrDate))
					{
						cells[inrDateColumnIndex] = inrDate.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						inrDateInvalidRows.Add(excelLineNumber);
						continue;
					}
					// GSTR-1 auto-population/ deletion upon cancellation date Validation
					string gstr1AutoPopulationValue = cells[gstr1AutoPopulationColumnIndex]?.Trim();
					if (DateTime.TryParseExact(gstr1AutoPopulationValue, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime gstr1Date))
					{
						cells[gstr1AutoPopulationColumnIndex] = gstr1Date.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						gstr1AutoPopulationInvalidRows.Add(excelLineNumber);
						continue;
					}




					dt.Rows.Add(cells);
				}

			}

			// Handle mismatch errors
			if (columnMismatchRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || supplierNameInvalidRows.Count > 0 || invoiceNumberInvalidRows.Count > 0 || invoiceDateInvalidRows.Count > 0 || eCommerceGstinInvalidRows.Count > 0 || inrDateInvalidRows.Count > 0 || gstr1AutoPopulationInvalidRows.Count > 0)
			{
				string error = "";
				if (columnMismatchRows.Count > 0)
					error += $"Column mismatch at rows: {string.Join(", ", columnMismatchRows)}\n";
				if (supplierGstinInvalidRows.Count > 0)
					error += $"Invalid Supplier GSTIN at rows: {string.Join(", ", supplierGstinInvalidRows)}\n";
				if (supplierNameInvalidRows.Count > 0)
					error += $"\n The Supplier Name field allows a maximum of 100 characters.Invalid Supplier Name at line(s) : {string.Join(", ", supplierNameInvalidRows)}.";
				if (invoiceNumberInvalidRows.Count > 0)
					error += $"\n The Invoice Number field allows a maximum of 25 characters.Invalid Invoice Number at line(s) : {string.Join(", ", invoiceNumberInvalidRows)}.";
				if (invoiceDateInvalidRows.Count > 0)
					error += $"\n Invalid Invoice Date at line(s) : {string.Join(", ", invoiceDateInvalidRows)}.";
				if (eCommerceGstinInvalidRows.Count > 0)
					error += $"\n Invalid E-Commerce GSTIN at line(s) : {string.Join(", ", eCommerceGstinInvalidRows)}.";
				if (inrDateInvalidRows.Count > 0)
					error += $"\n Invalid IRN Date at line(s) : {string.Join(", ", inrDateInvalidRows)}.";
				if (gstr1AutoPopulationInvalidRows.Count > 0)
					error += $"\n Invalid GSTR-1 auto-population/ deletion upon cancellation date at line(s) : {string.Join(", ", gstr1AutoPopulationInvalidRows)}.";

				throw new Exception("Failed to insert data: \n" + error);

			}
			return dt;
		}
		private DataTable ReadEInvoiceXLSFile(Stream stream, string sheetName, string gstin)
		{
			var SLEInvoiceColumns = _configuration["SLEInvoice"].Split(',').Select(x => x.Trim()).ToList();

			DataTable dt = new DataTable();
			int lineNumber = 0;
			//suppliergstin,supplier name,invoiceno, invoice date,E-Commerce GSTIN,IRN date ,GSTR-1 auto-population/ deletion upon cancellation date	

			List<int> columnMismatchRows = new List<int>();
			List<int> supplierGstinInvalidRows = new List<int>();
			List<int> supplierNameInvalidRows = new List<int>();
			List<int> invoiceNumberInvalidRows = new List<int>();
			List<int> invoiceDateInvalidRows = new List<int>();
			List<int> eCommerceGstinInvalidRows = new List<int>();
			List<int> inrDateInvalidRows = new List<int>();
			List<int> gstr1AutoPopulationInvalidRows = new List<int>();

			// Load the .xls workbook using NPOI
			HSSFWorkbook workbook = new HSSFWorkbook(stream);
			ISheet sheet = workbook.GetSheet(sheetName) ?? throw new Exception($"Sheet '{sheetName}' not found.");

			if (sheet.PhysicalNumberOfRows == 0)
				throw new Exception("Excel file is empty.");

			// Read header
			IRow headerRow = sheet.GetRow(0);
			int columnCount = headerRow.LastCellNum;
			string[] headers = new string[columnCount];

			for (int i = 0; i < columnCount; i++)
			{
				headers[i] = headerRow.GetCell(i)?.ToString().Trim() ?? $"Column{i}";
				dt.Columns.Add(headers[i], typeof(string));
			}
			// Index mapping and validation
			int supplierGstinIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[0]);
			int supplierNameColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[1]);
			int invoiceNumberColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[2]);
			int invoiceDateColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[3]);
			int eCommerceGstinColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[9]);
			int inrDateColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[17]);
			int gstr1AutoPopulationColumnIndex = FindEInvoiceColumnIndex(headers, SLEInvoiceColumns[19]);
			

			// Process rows
			for (int i = 1; i <= sheet.LastRowNum; i++)
			{
				IRow row = sheet.GetRow(i);
				if (row == null) continue;

				string[] cells = new string[columnCount];
				for (int j = 0; j < columnCount; j++)
				{
					ICell cell = row.GetCell(j);
					if (cell == null)
					{
						cells[j] = string.Empty;
						continue;
					}

					if ((j == invoiceDateColumnIndex || j == inrDateColumnIndex) && cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
					{
						DateTime DT = (DateTime)cell.DateCellValue;
						cells[j] = DT.ToString("dd-MM-yyyy hh:mm:ss tt"); // Match your allowed format
					}
					else
					{
						cells[j] = cell.ToString().Trim();
					}
				}


				int excelLineNumber = i + 1;

				// Column count check
				if (cells.Length != dt.Columns.Count)
				{
					columnMismatchRows.Add(excelLineNumber);
					continue;
				}
				// Supplier GSTIN Validation
				string supplierGstinValue = cells[supplierGstinIndex];
				if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
				{
					supplierGstinInvalidRows.Add(excelLineNumber);
					continue;
				}
				// Supplier Name Validation
				string supplierNameValue = cells[supplierNameColumnIndex]?.Trim();
				if (supplierNameValue.Length > 100)
				{
					supplierNameInvalidRows.Add(excelLineNumber);
					continue;
				}
				// Invoice Number Validation
				string invoiceNumberValue = cells[invoiceNumberColumnIndex]?.Trim();
				if (string.IsNullOrWhiteSpace(invoiceNumberValue) || invoiceNumberValue.Length > 25)
				{
					invoiceNumberInvalidRows.Add(excelLineNumber);
					continue;
				}
				// Invoice Date Validation
				string invoiceDateStr = cells[invoiceDateColumnIndex]?.Trim();
				string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
				if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
				{
					cells[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
				}
				else
				{
					invoiceDateInvalidRows.Add(excelLineNumber);
					continue;
				}
				// E-Commerce GSTIN Validation
				string eCommerceGstinValue = cells[eCommerceGstinColumnIndex]?.Trim();
				if (string.IsNullOrWhiteSpace(eCommerceGstinValue) || eCommerceGstinValue.Length != 15)
				{
					eCommerceGstinInvalidRows.Add(excelLineNumber);
					continue;
				}
				// IRN Date Validation
				string inrDateStr = cells[inrDateColumnIndex]?.Trim();
				if (DateTime.TryParseExact(inrDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime inrDate))
				{
					cells[inrDateColumnIndex] = inrDate.ToString("dd-MM-yyyy hh:mm:ss tt");
				}
				else
				{
					inrDateInvalidRows.Add(excelLineNumber);
					continue;
				}
				// GSTR-1 auto-population/ deletion upon cancellation date Validation
				string gstr1AutoPopulationValue = cells[gstr1AutoPopulationColumnIndex]?.Trim();
				if (DateTime.TryParseExact(gstr1AutoPopulationValue, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime gstr1Date))
				{
					cells[gstr1AutoPopulationColumnIndex] = gstr1Date.ToString("dd-MM-yyyy hh:mm:ss tt");
				}
				else
				{
					gstr1AutoPopulationInvalidRows.Add(excelLineNumber);
					continue;
				}





				// Add row to DataTable
				dt.Rows.Add(cells);
			}

			// Handle mismatch errors
			if (columnMismatchRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || supplierNameInvalidRows.Count > 0 || invoiceNumberInvalidRows.Count > 0 || invoiceDateInvalidRows.Count > 0 || eCommerceGstinInvalidRows.Count > 0 || inrDateInvalidRows.Count > 0 || gstr1AutoPopulationInvalidRows.Count > 0)
			{
				string error = "";
				if (columnMismatchRows.Count > 0)
					error += $"Column mismatch at rows: {string.Join(", ", columnMismatchRows)}\n";
				if (supplierGstinInvalidRows.Count > 0)
					error += $"Invalid Supplier GSTIN at rows: {string.Join(", ", supplierGstinInvalidRows)}\n";
				if (supplierNameInvalidRows.Count > 0)
					error += $"\n The Supplier Name field allows a maximum of 100 characters.Invalid Supplier Name at line(s) : {string.Join(", ", supplierNameInvalidRows)}.";
				if (invoiceNumberInvalidRows.Count > 0)
					error += $"\n The Invoice Number field allows a maximum of 25 characters.Invalid Invoice Number at line(s) : {string.Join(", ", invoiceNumberInvalidRows)}.";
				if (invoiceDateInvalidRows.Count > 0)
					error += $"\n Invalid Invoice Date at line(s) : {string.Join(", ", invoiceDateInvalidRows)}.";
				if (eCommerceGstinInvalidRows.Count > 0)
					error += $"\n Invalid E-Commerce GSTIN at line(s) : {string.Join(", ", eCommerceGstinInvalidRows)}.";
				if (inrDateInvalidRows.Count > 0)
					error += $"\n Invalid IRN Date at line(s) : {string.Join(", ", inrDateInvalidRows)}.";
				if (gstr1AutoPopulationInvalidRows.Count > 0)
					error += $"\n Invalid GSTR-1 auto-population/ deletion upon cancellation date at line(s) : {string.Join(", ", gstr1AutoPopulationInvalidRows)}.";

				throw new Exception("Failed to insert data: \n" + error);

			}
			return dt;
		}
		private bool ValidateEInvoiceColumnNames(DataTable dataTable)
		{
			var SLEInvoice = _configuration["SLEInvoice"];
			var _SLEInvoiceColumns = SLEInvoice.Split(',').Select(x => x.Trim()).ToList();

			var uploadedColumns = dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName.ToLower()).ToList();
			return _SLEInvoiceColumns.All(col => uploadedColumns.Contains(col.ToLower()));
		}
		private int FindEInvoiceColumnIndex(string[] headers, string expectedName)
		{
			int index = Array.FindIndex(headers, h => h.Equals(expectedName, StringComparison.OrdinalIgnoreCase));
			if (index == -1)
				throw new Exception($"The EInvoice file does not contain a '{expectedName}' column.");
			return index;
		}

		private DataTable ReadEWayBillCsvFile(Stream stream, string gstin)
		{
			var EWayBillColumns = _configuration["SLEWayBill"].Split(',').Select(x => x.Trim()).ToList();

			//foreach(var item in EWayBillColumns)
			//{
			//	Console.WriteLine($"EWayBillColumns : {item}");	
			//}

			DataTable dt = new DataTable();
			int lineNumber = 0;

			List<int> columnMismatchRows = new List<int>();

			List<int> ewbDateInvalidRows = new List<int>();
			List<int> invoiceNumberInvalidRows = new List<int>();
			List<int> invoiceDateInvalidRows = new List<int>();
			List<int> supplierGstinInvalidRows = new List<int>();
			List<int> validTillDateInvalidRows = new List<int>();

			using (var reader = new StreamReader(stream))
			using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(reader))
			{
				parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
				parser.SetDelimiters(",");
				parser.HasFieldsEnclosedInQuotes = true;

				// Read headers
				string[] headers = parser.ReadFields();
				lineNumber++;
				foreach (string header in headers)
				{
					dt.Columns.Add(header.Trim(), typeof(string));
					//Console.WriteLine($"Headers of file : {header}");
				}

				// Index mapping and validation
				int EwbDateColumnIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[1]);
				int invoiceNumberColumnIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[3]);
				int invoiceDateColumnIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[4]);
				int supplierGstinIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[6]);
				int validTillDateColumnIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[17]);

				// Read data rows
				while (!parser.EndOfData)
				{
					string[] fields = parser.ReadFields();
					lineNumber++;

					//Console.WriteLine($"fields.Length {fields.Length}");
					//Console.WriteLine($"dt.Columns.Count {dt.Columns.Count}");
					if (fields.Length != dt.Columns.Count)
					{
						columnMismatchRows.Add(lineNumber);
						continue;
					}
					//Ewb Date Validation
					string ewbDateStr = fields[EwbDateColumnIndex]?.Trim();
					string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
					if (DateTime.TryParseExact(ewbDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime ewbDate))
					{
						fields[EwbDateColumnIndex] = ewbDate.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						ewbDateInvalidRows.Add(lineNumber);
						continue;
					}
					// Invoice Number Validation
					string invoiceNumberValue = fields[invoiceNumberColumnIndex]?.Trim();
					if (string.IsNullOrWhiteSpace(invoiceNumberValue) || invoiceNumberValue.Length > 25)
					{
						invoiceNumberInvalidRows.Add(lineNumber);
						continue;
					}
					// Invoice Date Validation
					string invoiceDateStr = fields[invoiceDateColumnIndex]?.Trim();
					//Console.WriteLine($"Invoice Date: {invoiceDateStr}");
					if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
					{
						fields[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						invoiceDateInvalidRows.Add(lineNumber);
						continue;
					}
					// Supplier GSTIN Validation
					string supplierGstinValue = fields[supplierGstinIndex];
					if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
					{
						supplierGstinInvalidRows.Add(lineNumber);
						continue;
					}
					// Valid Till Date Validation
					string validTillDateStr = fields[validTillDateColumnIndex]?.Trim();
					//Console.WriteLine($"Valid Till Date: {validTillDateStr}");
					if (DateTime.TryParseExact(validTillDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime validTillDate))
					{
						fields[validTillDateColumnIndex] = validTillDate.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						validTillDateInvalidRows.Add(lineNumber);
						continue;
					}




					dt.Rows.Add(fields);
				}
			}
			// Handle mismatch errors
			if (columnMismatchRows.Count > 0 || ewbDateInvalidRows.Count > 0 || invoiceNumberInvalidRows.Count > 0 || invoiceDateInvalidRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || validTillDateInvalidRows.Count > 0)
			{
				string error = "";
				if (columnMismatchRows.Count > 0)
					error += $"Column mismatch at rows: {string.Join(", ", columnMismatchRows)}\n";
				if (ewbDateInvalidRows.Count > 0)
					error += $"Invalid EWB Date at rows: {string.Join(", ", ewbDateInvalidRows)}\n";
				if (invoiceNumberInvalidRows.Count > 0)
					error += $"\n The Invoice Number field allows a maximum of 25 characters.Invalid Invoice Number at line(s) : {string.Join(", ", invoiceNumberInvalidRows)}.";
				if (invoiceDateInvalidRows.Count > 0)
					error += $"\n Invalid Invoice Date at line(s) : {string.Join(", ", invoiceDateInvalidRows)}. Expected Formate : dd-MM-yyyy hh:mm:ss tt .";
				if (supplierGstinInvalidRows.Count > 0)
					error += $"Invalid Supplier GSTIN at rows: {string.Join(", ", supplierGstinInvalidRows)}\n";
				if (validTillDateInvalidRows.Count > 0)
					error += $"Invalid 'Valid Till Date' at rows: {string.Join(", ", validTillDateInvalidRows)}\n";

				throw new Exception("Failed to insert data file EWayBill :  \n" + error);

			}
			return dt;
		}
		private DataTable ReadEwayBillExcelFile(Stream stream, string sheetName, string gstin)
		{
			var EWayBillColumns = _configuration["SLEWayBill"].Split(',').Select(x => x.Trim()).ToList();

			DataTable dt = new DataTable();
			int lineNumber = 0;

			List<int> columnMismatchRows = new List<int>();

			List<int> ewbDateInvalidRows = new List<int>();
			List<int> invoiceNumberInvalidRows = new List<int>();
			List<int> invoiceDateInvalidRows = new List<int>();
			List<int> supplierGstinInvalidRows = new List<int>();
			List<int> validTillDateInvalidRows = new List<int>();

			using (var workbook = new XLWorkbook(stream))
			{
				var worksheet = workbook.Worksheet(sheetName);
				if (worksheet == null)
					throw new Exception($"Sheet '{sheetName}' not found.");
				var rows = worksheet.RowsUsed().ToList();

				if (rows.Count == 0)
					throw new Exception("Excel file is empty.");

				// Read header
				var headerRow = rows[0];
				var headers = headerRow.Cells().Select(c => c.GetString().Trim()).ToArray();

				foreach (var header in headers)
					dt.Columns.Add(header, typeof(string));


				// Index mapping and validation
				int EwbDateColumnIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[1]);
				int invoiceNumberColumnIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[3]);
				int invoiceDateColumnIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[4]);
				int supplierGstinIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[6]);
				int validTillDateColumnIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[17]);

				for (int i = 1; i < rows.Count; i++)
				{
					int excelLineNumber = i + 1; // Line number as seen in Excel (header is line 1)
					var row = rows[i];
					var cells = row.Cells().Select(c => c.GetString().Trim()).ToArray();
					// Column count check
					if (cells.Length != dt.Columns.Count)
					{
						columnMismatchRows.Add(excelLineNumber);
						continue;
					}
					//Ewb Date Validation
					string ewbDateStr = cells[EwbDateColumnIndex]?.Trim();
					string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
					if (DateTime.TryParseExact(ewbDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime ewbDate))
					{
						cells[EwbDateColumnIndex] = ewbDate.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						ewbDateInvalidRows.Add(excelLineNumber);
						continue;
					}
					// Invoice Number Validation
					string invoiceNumberValue = cells[invoiceNumberColumnIndex]?.Trim();
					if (string.IsNullOrWhiteSpace(invoiceNumberValue) || invoiceNumberValue.Length > 25)
					{
						invoiceNumberInvalidRows.Add(excelLineNumber);
						continue;
					}
					// Invoice Date Validation
					string invoiceDateStr = cells[invoiceDateColumnIndex]?.Trim();
					if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
					{
						cells[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						invoiceDateInvalidRows.Add(excelLineNumber);
						continue;
					}
					// Supplier GSTIN Validation
					string supplierGstinValue = cells[supplierGstinIndex];
					if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
					{
						supplierGstinInvalidRows.Add(excelLineNumber);
						continue;
					}
					// Valid Till Date Validation
					string validTillDateStr = cells[validTillDateColumnIndex]?.Trim();
					Console.WriteLine($"Valid Till Date: {validTillDateStr}");
					if (DateTime.TryParseExact(validTillDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime validTillDate))
					{
						cells[validTillDateColumnIndex] = validTillDate.ToString("dd-MM-yyyy hh:mm:ss tt");
					}
					else
					{
						validTillDateInvalidRows.Add(excelLineNumber);
						continue;
					}

					dt.Rows.Add(cells);


				}
			}
			if (columnMismatchRows.Count > 0 || ewbDateInvalidRows.Count > 0 || invoiceNumberInvalidRows.Count > 0 || invoiceDateInvalidRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || validTillDateInvalidRows.Count > 0)
			{
				string error = "";
				if (columnMismatchRows.Count > 0)
					error += $"Column mismatch at rows: {string.Join(", ", columnMismatchRows)}\n";
				if (ewbDateInvalidRows.Count > 0)
					error += $"Invalid EWB Date at rows: {string.Join(", ", ewbDateInvalidRows)}\n";
				if (invoiceNumberInvalidRows.Count > 0)
					error += $"\n The Invoice Number field allows a maximum of 25 characters.Invalid Invoice Number at line(s) : {string.Join(", ", invoiceNumberInvalidRows)}.";
				if (invoiceDateInvalidRows.Count > 0)
					error += $"\n Invalid Invoice Date at line(s) : {string.Join(", ", invoiceDateInvalidRows)}. Expected Formate : dd-MM-yyyy hh:mm:ss tt .";
				if (supplierGstinInvalidRows.Count > 0)
					error += $"Invalid Supplier GSTIN at rows: {string.Join(", ", supplierGstinInvalidRows)}\n";
				if (validTillDateInvalidRows.Count > 0)
					error += $"Invalid 'Valid Till Date' at rows: {string.Join(", ", validTillDateInvalidRows)}\n";

				throw new Exception("Failed to insert data of file EWayBill  \n" + error);

			}
			return dt;
		}
		private DataTable ReadEwayBillXLSFile(Stream stream, string sheetName, string gstin)
		{
			var EWayBillColumns = _configuration["SLEWayBill"].Split(',').Select(x => x.Trim()).ToList();

			DataTable dt = new DataTable();
			int lineNumber = 0;

			// ewb date,docno,docdate,togstin,Valid Till Date


			List<int> columnMismatchRows = new List<int>();

			List<int> ewbDateInvalidRows = new List<int>();
			List<int> invoiceNumberInvalidRows = new List<int>();
			List<int> invoiceDateInvalidRows = new List<int>();
			List<int> supplierGstinInvalidRows = new List<int>();
			List<int> validTillDateInvalidRows = new List<int>();

			// Load the .xls workbook using NPOI
			HSSFWorkbook workbook = new HSSFWorkbook(stream);
			ISheet sheet = workbook.GetSheet(sheetName) ?? throw new Exception($"Sheet '{sheetName}' not found.");

			if (sheet.PhysicalNumberOfRows == 0)
				throw new Exception("Excel file is empty.");

			// Read header
			IRow headerRow = sheet.GetRow(0);
			int columnCount = headerRow.LastCellNum;
			string[] headers = new string[columnCount];

			for (int i = 0; i < columnCount; i++)
			{
				headers[i] = headerRow.GetCell(i)?.ToString().Trim() ?? $"Column{i}";
				dt.Columns.Add(headers[i], typeof(string));
			}
            // Index mapping and validation
            //EWB No,EWB Date,Supply Type,Doc.No,Doc.Date,Doc.Type,
			//TO GSTIN,status,No of Items,Main HSN Code,Main HSN Desc,
			//Assessable Value,SGST Value,CGST Value,IGST Value,CESS Value,
			//Total Invoice Value,Valid Till Date,Gen.Mode
            int EwbDateColumnIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[1]);
			int invoiceNumberColumnIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[3]);
			int invoiceDateColumnIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[4]);
			int supplierGstinIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[6]);
			int validTillDateColumnIndex = FindEwayBillColumnIndex(headers, EWayBillColumns[17]);
			


			// Process rows
			for (int i = 1; i <= sheet.LastRowNum; i++)
			{
				IRow row = sheet.GetRow(i);
				if (row == null) continue;

				string[] cells = new string[columnCount];
				for (int j = 0; j < columnCount; j++)
				{
					ICell cell = row.GetCell(j);
					if (cell == null)
					{
						cells[j] = string.Empty;
						continue;
					}

					if ((j == EwbDateColumnIndex || j == invoiceDateColumnIndex || j == validTillDateColumnIndex) && cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
					{
						DateTime DT = (DateTime)cell.DateCellValue;
						cells[j] = DT.ToString("dd-MM-yyyy hh:mm:ss tt"); // Match your allowed format
					}
					else
					{
						cells[j] = cell.ToString().Trim();
					}
				}


				int excelLineNumber = i + 1;

				if (cells.Length != dt.Columns.Count)
				{
					columnMismatchRows.Add(excelLineNumber);
					continue;
				}
				//Ewb Date Validation
				string ewbDateStr = cells[EwbDateColumnIndex]?.Trim();
				string[] allowedFormats = _configuration["Date Format"].Split(',').Select(x => x.Trim()).ToArray();
				if (DateTime.TryParseExact(ewbDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime ewbDate))
				{
					cells[EwbDateColumnIndex] = ewbDate.ToString("dd-MM-yyyy hh:mm:ss tt");
				}
				else
				{
					ewbDateInvalidRows.Add(excelLineNumber);
					continue;
				}
				// Invoice Number Validation
				string invoiceNumberValue = cells[invoiceNumberColumnIndex]?.Trim();
				if (string.IsNullOrWhiteSpace(invoiceNumberValue) || invoiceNumberValue.Length > 25)
				{
					invoiceNumberInvalidRows.Add(excelLineNumber);
					continue;
				}
				// Invoice Date Validation
				string invoiceDateStr = cells[invoiceDateColumnIndex]?.Trim();
				if (DateTime.TryParseExact(invoiceDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invoiceDate))
				{
					cells[invoiceDateColumnIndex] = invoiceDate.ToString("dd-MM-yyyy hh:mm:ss tt");
				}
				else
				{
					invoiceDateInvalidRows.Add(excelLineNumber);
					continue;
				}
				// Supplier GSTIN Validation
				string supplierGstinValue = cells[supplierGstinIndex];
				if (string.IsNullOrWhiteSpace(supplierGstinValue) || supplierGstinValue.Length != 15)
				{
					supplierGstinInvalidRows.Add(excelLineNumber);
					continue;
				}
				// Valid Till Date Validation
				string validTillDateStr = cells[validTillDateColumnIndex]?.Trim();
				//Console.WriteLine($"Valid Till Date: {validTillDateStr}");
				if (DateTime.TryParseExact(validTillDateStr, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime validTillDate))
				{
					cells[validTillDateColumnIndex] = validTillDate.ToString("dd-MM-yyyy hh:mm:ss tt");
				}
				else
				{
					validTillDateInvalidRows.Add(excelLineNumber);
					continue;
				}

				dt.Rows.Add(cells);

			}


			if (columnMismatchRows.Count > 0 || ewbDateInvalidRows.Count > 0 || invoiceNumberInvalidRows.Count > 0 || invoiceDateInvalidRows.Count > 0 || supplierGstinInvalidRows.Count > 0 || validTillDateInvalidRows.Count > 0)
			{
				string error = "";
				if (columnMismatchRows.Count > 0)
					error += $"Column mismatch at rows: {string.Join(", ", columnMismatchRows)}\n";
				if (ewbDateInvalidRows.Count > 0)
					error += $"Invalid EWB Date at rows: {string.Join(", ", ewbDateInvalidRows)}\n";
				if (invoiceNumberInvalidRows.Count > 0)
					error += $"\n The Invoice Number field allows a maximum of 25 characters.Invalid Invoice Number at line(s) : {string.Join(", ", invoiceNumberInvalidRows)}.";
				if (invoiceDateInvalidRows.Count > 0)
					error += $"\n Invalid Invoice Date at line(s) : {string.Join(", ", invoiceDateInvalidRows)}. Expected Formate : dd-MM-yyyy hh:mm:ss tt .";
				if (supplierGstinInvalidRows.Count > 0)
					error += $"Invalid Supplier GSTIN at rows: {string.Join(", ", supplierGstinInvalidRows)}\n";
				if (validTillDateInvalidRows.Count > 0)
					error += $"Invalid 'Valid Till Date' at rows: {string.Join(", ", validTillDateInvalidRows)}\n";

				throw new Exception("Failed to insert data of file EWayBill  \n" + error);

			}
			return dt;
		}
		private bool ValidateEWayBillColumnNames(DataTable dataTable)
		{
			var EWayBill = _configuration["SLEWayBill"];
			var EWayBillColumns = EWayBill.Split(',').Select(x => x.Trim()).ToList();

			var uploadedColumns = dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName.ToLower()).ToList();
			return EWayBillColumns.All(col => uploadedColumns.Contains(col.ToLower()));
		}
		private int FindEwayBillColumnIndex(string[] headers, string expectedName)
		{
			int index = Array.FindIndex(headers, h => h.Equals(expectedName, StringComparison.OrdinalIgnoreCase));
			if (index == -1)
				throw new Exception($"The EwayBill file does not contain a '{expectedName}' column.");
			return index;
		}

		#endregion

		#region SalesLedgerCurrentRequestsAPI-Master
		public async Task<IActionResult> SalesLedgerCurrentRequestsAPI_Master(DateTime? fromdate, DateTime? todate)
		{
			DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
			DateTime toDateTime = todate ?? DateTime.Now; // Default to today
			//await TjCaptions("OpenTasks");

			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act6";
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                return View("~/Views/Admin/SalesLedgerCurrentRequests/SalesLedgerCurrentRequestsAPI_Master.cshtml");
            }

            //         var email = MySession.Current.Email; // Get the email from session
            //var clients = await _userBusiness.GetAdminClients(email);
            //string[] gstinArray = clients.Select(c => c.ClientGSTIN).ToArray();
            //var Tickets = await _sLTicketsBusiness.GetClientsOpenTicketsStatusAsync(gstinArray, fromDateTime, toDateTime);

            var Tickets = await _sLTicketsBusiness.GetOpenTicketsStatusAsync(MySession.Current.gstin, fromDateTime, toDateTime); //(ViewBag.UserGSTIN); // Fetch pending tickets based on user_gstin

            ViewBag.Tickets = Tickets;

			return View("~/Views/Admin/SalesLedgerCurrentRequests/SalesLedgerCurrentRequestsAPI_Master.cshtml");
		}

		public async Task<IActionResult> CompareSalesLedgerGSTAPI_Master(string ticketnumber, string ClientGSTIN, string fromDate, string toDate)
		{
			// _logger.LogInformation($"Ticket Number: {ticketnumber}");
			//based on this ticket number fetch data from Purchase Ticket table and store
			var Ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);

			ViewBag.Ticket = Ticket;
			ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            ViewBag.fromDate = fromDate;
            ViewBag.toDate = toDate;
            return View("~/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTAPI_Master.cshtml");
		}

        [HttpPost]
        public async Task<IActionResult> CompareSalesLedgerAPI_Master(string ticketnumber, string ClientGSTIN)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            string sessionId = HttpContext.Session.Id;
            var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            ViewBag.Ticket = Ticket;

            string UserName = "";
            try
            {
                var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);
                UserName = UserAPIData.GstPortalUsername;
            }
            catch
            {
                string errorMessage = "Error fetching user API data. Please update user API data.";
                return Json(new { failure = true, message = errorMessage });
            }

            string Email = _configuration["EInvoiceMatserAPI:email"];

            try
            {
                var authTokenData = await _gSTR2DataBusiness.GetTokenDataAsync(ClientGSTIN);

                bool isTokenExpired = authTokenData == null ||
                                      string.IsNullOrEmpty(authTokenData.AuthToken) ||
                                      (authTokenData.AuthTokenCreatedDatetime?.AddSeconds(Convert.ToDouble(authTokenData.Expiry)) < DateTime.Now);

                if (!isTokenExpired)
                {
                    return Json(new { success = true, message = "Token is valid. Continue." });
                }

                #region Request for authentation  1st Api call

                // Parameters
                string Parameters = $"email={Email}";

                // Add required headers
                string userName = UserName;
                string stateCd = ClientGSTIN.Substring(0, 2);
                string ipAddress = _configuration["EInvoiceMatserAPI:ipAddress"];
                string clientId = _configuration["EInvoiceMatserAPI:ClientId"];
                string clientSecret = _configuration["EInvoiceMatserAPI:ClientSecret"];

                var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Add("gst_username", userName);
                httpClient.DefaultRequestHeaders.Add("state_cd", stateCd);
                httpClient.DefaultRequestHeaders.Add("ip_address", ipAddress);
                httpClient.DefaultRequestHeaders.Add("client_id", clientId);
                httpClient.DefaultRequestHeaders.Add("client_secret", clientSecret);

                // Body
                var RequestBody = new { };

                // Url
                string baseUrl = _configuration["EInvoiceMatserAPI:BaseUrl"] ?? throw new InvalidOperationException("BaseUrl for EInvoiceMatserAPI is not configured in ApiSettings");
                string otpRequestEndpoint = _configuration["EInvoiceMatserAPI:OtpRequest"] ?? throw new InvalidOperationException("OtpRequest for EInvoiceMatserAPI is not configured in ApiSettings");
                string apiUrl1 = $"{baseUrl}{otpRequestEndpoint}?{Parameters}";


                // API call
                var response = await httpClient.GetAsync(apiUrl1);
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);

                // save API data to database
                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                var headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketnumber;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl1;
                    ApiData.RequestParameters = Parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _sLEInvoiceBusiness.saveEInvoiceAPIsData(ApiData);


                // Validate response 

                Console.WriteLine($"Response 1 : {responseData}");
                Console.WriteLine($"Result 1 : {result}");
                //Console.ReadKey();

                //Result 1  {
                //	"status_cd": "1",
                //	"status_desc": "user name exists",
                //	"header": {
                //		"gst_username": "MH_NT2.1642",
                //		"state_cd": "27",
                //		"ip_address": "14.98.237.54",
                //		"client_id": "GSTS36731818-d2d3-4036-816c-d3e559eb1dde",
                //		"client_secret": "GSTSa012637b-1daa-43be-aa1f-85d83ce132fb",
                //		"cache-control": "no-cache",
                //		"postman-token": "446ef596-b536-4e9a-8f71-887af9324664",
                //		"txn": "5dcd07ed2a714698b9e74ccea93bf775"
                //	}
                //}
                string txn = null;

                if (result["header"] != null && result["header"]["txn"] != null)
                {
                    txn = result["header"]["txn"].ToString();
                    return Json(new { success = false, askForOtp = true, txn = txn });
                }
                if (result["status_cd"].ToString() != "1" || txn == null)
                {
                    string errorMessage = $"{result["error"]["message"]}";
                    return Json(new
                    {
                        failure = true,
                        message = $"E-Invoice OTP Request API Call - Failed due to :{errorMessage}" //Msg from result 
                    });
                }

                return Json(new { failure = true, message = "E-Invoice OTP Request API Call - Failed ." });
                #endregion
            }
            catch (Exception ex)
            {
                string errorMessage = $"Error : {ex.Message}";
                return Json(new { failure = true, message = errorMessage });
            }
        }

        [HttpPost]
        public async Task<IActionResult> SLSubmitOtpAndContinue_Master(string ClientGSTIN, string ticketNo, string otp, string txn)
        {
            string sessionId = HttpContext.Session.Id;
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);

            string UserName = UserAPIData.GstPortalUsername;
            string Email = _configuration["EInvoiceMatserAPI:email"];

            try
            {
                #region 2nd Api call

                // Parameters
                string Parameters = $"email={Email}&otp={otp}";

                // Add required headers
                string stateCd = ClientGSTIN.Substring(0, 2);
                string ipAddress = _configuration["EInvoiceMatserAPI:ipAddress"];
                string clientId = _configuration["EInvoiceMatserAPI:ClientId"];
                string clientSecret = _configuration["EInvoiceMatserAPI:ClientSecret"];

                var httpClient = _httpClientFactory.CreateClient();

                httpClient.DefaultRequestHeaders.Add("gst_username", UserName);
                httpClient.DefaultRequestHeaders.Add("state_cd", stateCd);
                httpClient.DefaultRequestHeaders.Add("ip_address", ipAddress);
                httpClient.DefaultRequestHeaders.Add("txn", txn);
                httpClient.DefaultRequestHeaders.Add("client_id", clientId);
                httpClient.DefaultRequestHeaders.Add("client_secret", clientSecret);

                // Body
                var RequestBody = new { };

                // Url
                string baseUrl = _configuration["EInvoiceMatserAPI:BaseUrl"] ?? throw new InvalidOperationException("BaseUrl for EInvoiceMatserAPI is not configured in ApiSettings");
                string authTokenEndpoint = _configuration["EInvoiceMatserAPI:AuthToken"] ?? throw new InvalidOperationException("AuthToken for EInvoiceMatserAPI is not configured in ApiSettings");
                string apiUrl = $"{baseUrl}{authTokenEndpoint}?{Parameters}";

                // API call
                var response = await httpClient.GetAsync(apiUrl);
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);

                // save API data to database
                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                var headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketNo;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = Parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _sLEInvoiceBusiness.saveEInvoiceAPIsData(ApiData);

                // Validate response 

                Console.WriteLine($"Response 2 : {responseData}");
                Console.WriteLine($"Result 2 : {result}");
                //Console.ReadKey();

                //Result 2 {
                //	"status_cd": "1",
                //	"status_desc": "If authentication succeeds",
                //	"header": {
                //		"gst_username": "MH_NT2.1642",
                //		"state_cd": "27",
                //		"ip_address": "14.98.237.54",
                //		"client_id": "GSTS36731818-d2d3-4036-816c-d3e559eb1dde",
                //		"client_secret": "GSTSa012637b-1daa-43be-aa1f-85d83ce132fb",
                //		"txn": "5dcd07ed2a714698b9e74ccea93bf775",
                //		"cache-control": "no-cache",
                //		"postman-token": "2dbb2361-fa6d-40f4-b8df-925b17443db5"
                //	}
                //}
                if (result["status_cd"].ToString() == "1")
                {
                    await _gSTR2DataBusiness.saveTokenData(new GSTR2TokenDataModel
                    {
                        ClientGstin = ClientGSTIN,
                        RequestNumber = ticketNo,
                        UserName = UserName,
                        XAppKey = txn,
                        OTP = otp,
                        AuthToken = "",
                        Expiry = "",
                        SEK = ""
                    });
                    return Json(new { success = true, message = "OTP Verified. Token saved." });
                }

                //Result 2 : {                            // Invalid otp response 
                //    "status_cd": "0",
                //    "status_desc": "If authentication fails",
                //    "error": {
                //        "message": "Invalid Session",
                //        "error_cd": "AUTH4033"
                //    },
                //    "header": {
                //        "gst_username": "MH_NT2.1642",
                //        "state_cd": "27",
                //        "ip_address": "14.98.237.54",
                //        "client_id": "GSTS36731818-d2d3-4036-816c-d3e559eb1dde",
                //        "client_secret": "GSTSa012637b-1daa-43be-aa1f-85d83ce132fb",
                //        "txn": "5dcd07ed2a714698b9e74ccea93bf775",
                //        "cache-control": "no-cache",
                //        "postman-token": "480c5a42-e634-482f-933c-23219ddf24b5"
                //    }
                //}
                if (result["status_cd"].ToString() == "0" && result["error"]["error_cd"].ToString() == "AUTH4033")
                {
                    return Json(new { success = false, askAgain = true, message = "Invalid OTP. Please enter again." });
                }

                if (result["status_cd"].ToString() != "1")
                {
                    string errorMessage = $"E-Invoice Auth Token API Call - Failed due to : {result["error"]["message"]}";
                    return Json(new { failure = true, message = errorMessage });
                }

                return Json(new { failure = true, message = "E Invoice Auth Token API Call - Failed." });

                #endregion
            }
            catch (Exception ex)
            {
                string errorMessage = $"Error: {ex.Message}";
                return Json(new { failure = true, message = errorMessage });
            }
        }

        [HttpPost]
        public async Task<IActionResult> SLContinueWith4thApi_Master(string ticketnumber, string ClientGSTIN)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            string sessionId = HttpContext.Session.Id;
            var Ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            ViewBag.Ticket = Ticket;

            string Txn_Period = Ticket.Period;
            DateTime parsedDate;
            string period = "";
            // Parse the input string using the correct format
            if (DateTime.TryParseExact(Txn_Period, "MMM-yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
            {
                period = parsedDate.ToString("MMyyyy");
            }

            var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);

            string UserName = UserAPIData.GstPortalUsername;
            string Email = _configuration["EInvoiceMatserAPI:email"];
            string SupplyType = _configuration["EInvoiceMatserAPI:irnjsonsAPIsuptyp"];

            var authTokenData = await _gSTR2DataBusiness.GetTokenDataAsync(ClientGSTIN);
            try
            {
                #region 4th Api call

                // Parameters
                string Parameters = $"email={Email}&gstin={ClientGSTIN}&suptyp={SupplyType}&rtnprd={period}";

                // Add required headers
                string stateCd = ClientGSTIN.Substring(0, 2);
                string ipAddress = _configuration["EInvoiceMatserAPI:ipAddress"];
                string clientId = _configuration["EInvoiceMatserAPI:ClientId"];
                string clientSecret = _configuration["EInvoiceMatserAPI:ClientSecret"];

                var httpClient = _httpClientFactory.CreateClient();

                httpClient.DefaultRequestHeaders.Add("gst_username", UserName);
                httpClient.DefaultRequestHeaders.Add("state_cd", stateCd);
                httpClient.DefaultRequestHeaders.Add("ip_address", ipAddress);
                httpClient.DefaultRequestHeaders.Add("txn", authTokenData.XAppKey);
                httpClient.DefaultRequestHeaders.Add("client_id", clientId);
                httpClient.DefaultRequestHeaders.Add("client_secret", clientSecret);

                // Body
                var RequestBody = new { };

                // Url
                string baseUrl = _configuration["EInvoiceMatserAPI:BaseUrl"] ?? throw new InvalidOperationException("BaseUrl for EInvoiceMatserAPI is not configured in ApiSettings");
                string irnjsons = _configuration["EInvoiceMatserAPI:irnjsons"] ?? throw new InvalidOperationException("PortalData endpoint for EInvoiceMatserAPI is not configured in ApiSettings");
                string apiUrl = $"{baseUrl}{irnjsons}?{Parameters}";

                // API call
                var response = await httpClient.GetAsync(apiUrl); // ✅ Corrected: GET, not POST
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);

                // save API data to database

                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                var headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketnumber;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = Parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _sLEInvoiceBusiness.saveEInvoiceAPIsData(ApiData);


                // Validate response 
                Console.WriteLine($"Response 4 : {responseData}");
                Console.WriteLine($"Result 4 : {result}");
                //Console.ReadKey();

                //DataTable dataTable;
                //try
                //{
                //    dataTable = SL_Master_EInvoiceJsontoDataTable(responseData, ClientGSTIN);
                //    await _sLEInvoiceBusiness.SaveEInvoiceDataAsync(dataTable, ticketnumber, ClientGSTIN);
                //    return Json(new { goToCompare = true });
                //}
                //catch (Exception ex)
                //{
                //    string errorMessage = $"Error: {ex.Message}";
                //    return Json(new { failure = true, message = errorMessage });
                //}

                //Result 3 : {
                //    "data": {
                //        "est": "30",
                //		  "token": "d889e2328ec949aabf9b2be90aabba7c"
                //    },
                //	"status_cd": "2",
                //	"status_desc": "GSTR request succeeds",
                //	"header": {
                //        "gst_username": "MH_NT2.1642",
                //		"state_cd": "27",
                //		"ip_address": "14.98.237.54",
                //		"client_id": "GSTS36731818-d2d3-4036-816c-d3e559eb1dde",
                //		"client_secret": "GSTSa012637b-1daa-43be-aa1f-85d83ce132fb",
                //		"txn": "5dcd07ed2a714698b9e74ccea93bf775",
                //		"cache-control": "no-cache",
                //		"postman-token": "cee1bc5e-8e6b-451f-b288-d561dc0781e1",
                //		"gstin": "27AAGCB1286Q2Z3"
                //    }
                //}
                string token = null;
                // check if "data" node exists
                if (result["data"] != null && result["data"]["token"] != null)
                {
                    token = result["data"]["token"].ToString();
                    //return Json(new { goToCompare = true });
                    return Json(new { success = true, token = token });
                }

                //Result 5 : {
                //              "status_cd": "0",
                //              "status_desc": "GSTR request failed",
                //              "error": {
                //                "message": "File generation is in progress",
                //                "error_cd": "EINV30109"
                //              },
                //              "header": {
                //                "gst_username": "MH_NT2.1641",
                //                "state_cd": "27",
                //                "ip_address": "14.98.237.54",
                //                "txn": "6581dc1df79247838c1b7683081a876e",
                //                "client_id": "GSTS36731818-d2d3-4036-816c-d3e559eb1dde",
                //                "client_secret": "GSTSa012637b-1daa-43be-aa1f-85d83ce132fb",
                //                "traceparent": "00-8d642bc55fac3456ab72412df2bf4dec-76af85063d80f8a8-00",
                //                "ret_period": "032025",
                //                "gstin": "27AAGCB1286Q1Z4"
                //              }
                //            }

                if (result["status_cd"]?.ToString() == "0" && result["error"]["error_cd"]?.ToString() == "EINV30109")
                {
                    return Json(new { message = "File generation" });
                }

                //Result 3 : {
                //			  "status_cd": "0",
                //			  "status_desc": "GSTR request failed",
                //			  "error": {
                //				"message": "No document found for the provided Inputs",
                //				"error_cd": "RET13509"
                //			  },

                if (result["status_cd"]?.ToString() != "1" || string.IsNullOrEmpty(token))
                {
                    string errorMessage = $"E-Invoice irnjsons API Call - Failed due to  : {result["error"]["message"]}";
                    return Json(new { failure = true, message = errorMessage });
                }



                return Json(new { failure = true, message = "E-Invoice irnjsons API Call - Failed." });

                #endregion

            }
            catch (Exception ex)
            {
                string errorMessage = $"Error: {ex.Message}";
                return Json(new { failure = true, message = errorMessage });
            }
        }

        [HttpPost]
        public async Task<IActionResult> SLContinueWithFinalApi_Master(string ticketnumber, string ClientGSTIN, string token5)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            string sessionId = HttpContext.Session.Id;
            var ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            ViewBag.Ticket = ticket;

            var authTokenData = await _gSTR2DataBusiness.GetTokenDataAsync(ClientGSTIN);

            DataTable EInvoiceDataTable = new DataTable();
            // Read headers
            var EInvoiceColumns = _configuration["SLEInvoice"].Split(',').Select(x => x.Trim()).ToList();
            foreach (string header in EInvoiceColumns)
            {
                EInvoiceDataTable.Columns.Add(header.Trim(), typeof(string));
            }

            string txnPeriod = ticket.Period;
            string formattedPeriod = DateTime.ParseExact(txnPeriod, "MMM-yy", CultureInfo.InvariantCulture).ToString("MMyyyy");

            var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);
            string UserName = UserAPIData.GstPortalUsername;

            string Email = _configuration["EInvoiceMatserAPI:email"];

            #region Final API Call  for Einvoice
            try
            {
                #region  5th API call to get data

                // Parameters
                string parameter = $"email={Email}&gstin={ClientGSTIN}&rtnprd={formattedPeriod}&token={token5}";

                // Headers
                string stateCd = ClientGSTIN.Substring(0, 2);
                string ipAddress = _configuration["EInvoiceMatserAPI:ipAddress"];
                string clientId = _configuration["EInvoiceMatserAPI:ClientId"];
                string clientSecret = _configuration["EInvoiceMatserAPI:ClientSecret"];

                var httpClient = _httpClientFactory.CreateClient();

                httpClient.DefaultRequestHeaders.Add("gst_username", UserName);
                httpClient.DefaultRequestHeaders.Add("state_cd", stateCd);
                httpClient.DefaultRequestHeaders.Add("ip_address", ipAddress);
                httpClient.DefaultRequestHeaders.Add("txn", authTokenData.XAppKey);
                httpClient.DefaultRequestHeaders.Add("client_id", clientId);
                httpClient.DefaultRequestHeaders.Add("client_secret", clientSecret);

                // Body
                var RequestBody5 = new
                {

                };
                var content5 = new StringContent(JsonConvert.SerializeObject(RequestBody5), Encoding.UTF8, "application/json");

                // Url
                string baseUrl5 = _configuration["EInvoiceMatserAPI:BaseUrl"];
                string AuthUrl5 = _configuration["EInvoiceMatserAPI:filedtl"];
                string apiUrl5 = $"{baseUrl5}{AuthUrl5}?{parameter}";

                // Api call - Response
                var response5 = await httpClient.GetAsync(apiUrl5);
                var responseData5 = await response5.Content.ReadAsStringAsync();
                var result5 = JsonConvert.DeserializeObject<JObject>(responseData5);

                // save API data to database
                var headersDict5 = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                string headersJson5 = System.Text.Json.JsonSerializer.Serialize(headersDict5);
                APIsDataModel ApiData5 = new APIsDataModel
                {
                    ClientGstin = ClientGSTIN,
                    RequestNumber = ticketnumber,
                    SessionID = sessionId,
                    RequestURL = apiUrl5,
                    RequestParameters = parameter,
                    RequestHeaders = headersJson5,
                    RequestBody = JsonConvert.SerializeObject(RequestBody5),
                    Response = responseData5,
                    ResponseCode = $"{(int)response5.StatusCode} {response5.StatusCode}",
                    Status = getStatus((int)response5.StatusCode)
                };
                await _sLEInvoiceBusiness.saveEInvoiceAPIsData(ApiData5);

                Console.WriteLine($"Response 5 : {response5}");
                Console.WriteLine($"Result 5 : {result5}");
                //Console.ReadKey();

                //Result 5 : {
                //              "status_cd": "0",
                //              "status_desc": "GSTR request failed",
                //              "error": {
                //                "message": "File generation is in progress",
                //                "error_cd": "EINV30109"
                //              },
                //              "header": {
                //                "gst_username": "MH_NT2.1641",
                //                "state_cd": "27",
                //                "ip_address": "14.98.237.54",
                //                "txn": "6581dc1df79247838c1b7683081a876e",
                //                "client_id": "GSTS36731818-d2d3-4036-816c-d3e559eb1dde",
                //                "client_secret": "GSTSa012637b-1daa-43be-aa1f-85d83ce132fb",
                //                "traceparent": "00-8d642bc55fac3456ab72412df2bf4dec-76af85063d80f8a8-00",
                //                "ret_period": "032025",
                //                "gstin": "27AAGCB1286Q1Z4"
                //              }
                //            }

                if (result5["status_cd"]?.ToString() == "0" && result5["error"]["error_cd"]?.ToString() == "EINV30109")
                {
                    return Json(new { message = "File generation", token = token5 });
                }

                //Result 5 : {
                //              "status_cd": "0",
                //              "status_desc": "GSTR request failed",
                //              "error": {
                //                "message": "No Record found for the provided Inputs",
                //                "error_cd": "EINV30108"
                //              },
                //              "header": {
                //                "gst_username": "MH_NT2.1641",
                //                "state_cd": "27",
                //                "ip_address": "14.98.237.54",
                //                "txn": "f2692008737e43399d4c658095ca9bc7",
                //                "client_id": "GSTS36731818-d2d3-4036-816c-d3e559eb1dde",
                //                "client_secret": "GSTSa012637b-1daa-43be-aa1f-85d83ce132fb",
                //                "traceparent": "00-6d4efca1fab85c6eedb28ae4ab31eb25-8e827d776b5166b3-00",
                //                "ret_period": "032025",
                //                "gstin": "27AAGCB1286Q1Z4"
                //              }
                //            }

                if (result5["status_cd"]?.ToString() == "0" && result5["error"]["error_cd"]?.ToString() == "EINV30108")
                {
                    string errorMessage = $"E-Invoice GET FILEDETL API call - Failed due to : {result5["error"]["message"]?.ToString()}";
                    return Json(new { success = true });
                }

                //Console.WriteLine($"1 - {result5["status_code"]?.ToString() == "0"} , 2 - {result5["urls"] != null}");

                //Result 5 : {
                //			  "ek": "yNRzdodhcNN9Alqmsr9CcQ5S8QOw4yTWcpksee9sQHc=",
                //            "urls": [
                //		            {
                //					  "ic": 1,
                //                    "ul": "https://files.gst.gov.in/einvdownloads/05022025/JSON/2c5847edad8642e4b42a794142fc63fd/27AAFCI5032C1ZZ_122024_Received_1.tar.gz?md5=Sa5nyD8SE6q5l2EMZmNmRg&expires=1738915665",
                //                    "hash": "6c1be18b84bff4ae0c61c719de7c5c1a9a6e141888cb8bd140de37bcb9382c6d"
                //		            }
                //                    ],
                //           "fc": 1,
                //           "nextAvailable": "2025-02-06 13:37:12"
                //          }
                if (result5["urls"] != null)
            {
                string ek = result5["ek"]?.ToString();
                int fileCount = (int)result5["fc"];
                JArray URLs = (JArray)result5["urls"];
                string[] urls = URLs.Select(u => u["ul"]?.ToString()).ToArray();

                foreach (string url in urls)
                {
                    //Console.WriteLine($"URL: {url}");
                    // Step 1: Download tar.gz file
                    using HttpClient client = new HttpClient();
                    byte[] fileBytes = await client.GetByteArrayAsync(url);

                    // Step 2: Save it temporarily
                    string tempTarGzPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".tar.gz");
                    await System.IO.File.WriteAllBytesAsync(tempTarGzPath, fileBytes);

                    // Step 3: Extract .tar.gz
                    string extractedFolderPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                    Directory.CreateDirectory(extractedFolderPath);

                    // First extract .gz to .tar
                    string tempTarPath = Path.ChangeExtension(tempTarGzPath, ".tar");
                    using (FileStream originalFileStream = new FileStream(tempTarGzPath, FileMode.Open, FileAccess.Read))
                    using (FileStream decompressedFileStream = new FileStream(tempTarPath, FileMode.Create))
                    using (GZipStream decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress))
                    {
                        decompressionStream.CopyTo(decompressedFileStream);
                    }

                    // Now extract .tar file
                    using (var archive = TarArchive.Open(tempTarPath))
                    {
                        foreach (var entry in archive.Entries.Where(entry => !entry.IsDirectory))
                        {
                            string extractedFilePath = Path.Combine(extractedFolderPath, entry.Key);
                            entry.WriteToFile(extractedFilePath);

                            // Step 4: Read JSON file content
                            string jsonContent = await System.IO.File.ReadAllTextAsync(extractedFilePath);

                            // Step 5: Pass to DecryptEinvoiceResponse
                            EInvoiceDataTable = DecryptEinvoiceResponse(jsonContent, ek, ClientGSTIN, EInvoiceDataTable);
                        }
                    }

                    // Optional: Clean up temp files
                    System.IO.File.Delete(tempTarGzPath);
                    System.IO.File.Delete(tempTarPath);
                    Directory.Delete(extractedFolderPath, true);

                    //Console.WriteLine("Hi");

                }

                // Save the data to database
                await _sLEInvoiceBusiness.SaveEInvoiceDataAsync(EInvoiceDataTable, ticketnumber, ClientGSTIN);

                return Json(new { success = true, ticketno = ticketnumber, gstin = ClientGSTIN });

            }

                //Result 5 : {
                //           "status_code": 0,
                //           "message": "File generation is in progress"
                //           }
                if (result5["status_code"]?.ToString() == "0")
                {
                    string message = result5["message"]?.ToString();

                    if (!string.IsNullOrEmpty(message) && message.Contains("File generation"))
                    {

                        return Json(new { message = "File Generation", token = token5 });
                        //return View("/Views/Admin/CompareGstFiles/CompareGSTAPI_IRIS.cshtml");
                    }
                    else
                    {
                        return Json(new
                        {
                            error = true,
                            message = $"E-Invoice GET FILEDETL API Call Failed - Response - \"{message}\"",
                        });
                        //throw new Exception($"5th API Call - Failed due to {message}");
                    }
                }

                else if (result5["error_code"]?.ToString() == "500")
                {
                    string message = result5["message"]?.ToString();
                    message += result5["description"]?.ToString();
                    return Json(new
                    {
                        error = true,
                        message = message,
                    });
                }

                return Json(new { failure = true, message = "E-Invoice GET FILEDETL API Call - Failed." });

                #endregion

            }
            catch (Exception ex)
            {
                string errorMessage = $"Error: {ex.Message}";
                return Json(new { failure = true, message = errorMessage });
            }
            #endregion
        }

        public async Task<IActionResult> CompareSLDataTables_Master(string ticketnumber, string ClientGSTIN)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            string sessionId = HttpContext.Session.Id;
            var Ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            ViewBag.Ticket = Ticket;

            #region  E Way BIll APIS

            var period = Ticket.Period;
            List<string> dates = new List<string>();
            string formate = "dd/MM/yyyy";
            // Parse the month and year
            DateTime startDate = DateTime.ParseExact("01-" + period, "dd-MMM-yy", null); //DateTime.ParseExact("01-Apr-25", "dd-MMM-yy", null) ;// 
            int daysInMonth = DateTime.DaysInMonth(startDate.Year, startDate.Month);
            // Loop through each day in the month
            for (int day = 1; day <= daysInMonth; day++)
            {
                DateTime currentDate = new DateTime(startDate.Year, startDate.Month, day);
                string formattedDate = currentDate.ToString(formate, CultureInfo.InvariantCulture);

                //string formattedDate = currentDate.ToString(formate);
                dates.Add(formattedDate);
            }
            List<string> otherdays = await _sLDataBusiness.GetOtherDates(ticketnumber, ClientGSTIN, period, formate);
            dates.AddRange(otherdays);


            DataTable SLEWayBillData = new DataTable();
            // Read headers
            var Ewaybill = _configuration["SLEWayBill"];
            var ewaybillColumns = Ewaybill.Split(',').Select(x => x.Trim()).ToList();
            foreach (string header in ewaybillColumns)
            {
                SLEWayBillData.Columns.Add(header.Trim(), typeof(string));
            }

            string apiUserName = "";
            string apiPassword = "";
            try
            {
                var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);
                apiUserName = UserAPIData.APIIRISUsername;
                apiPassword = UserAPIData.APIIRISPassword;
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = "Error fetching user API data. Please update user API data.";
                return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTAPI_Master.cshtml");
            }

            string tempClientGSTIN = ClientGSTIN;
            if (ClientGSTIN == "33AAGCB1286Q2ZA" || ClientGSTIN == "27AAGCB1286Q1Z4" || ClientGSTIN == "33AAGCB1286Q1ZB" || ClientGSTIN == "27AAGCB1286Q2Z3")
            {
                ClientGSTIN = "29AAGCB1286Q000";
            }

            string Email = _configuration["EWayBillMatserAPI:email"];

            try
            {
                #region  1st API call
                // Parameters
                string Parameters = $"email={Email}&username={apiUserName}&password={apiPassword}";

                // Headers
                string clientId = _configuration["EWayBillMatserAPI:ClientId"];
                string clientSecret = _configuration["EWayBillMatserAPI:ClientSecret"];
                string ipAddress = _configuration["EWayBillMatserAPI:ipAddress"];
                string accept = _configuration["EWayBillMatserAPI:Accept"];

                var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Add("client_id", clientId);
                httpClient.DefaultRequestHeaders.Add("client_secret", clientSecret);
                httpClient.DefaultRequestHeaders.Add("gstin", ClientGSTIN);
                httpClient.DefaultRequestHeaders.Add("ip_address", ipAddress);
                httpClient.DefaultRequestHeaders.Add("Accept", accept);
                // Body
                var RequestBody = new { };

                // Url
                string baseUrl = _configuration["EWayBillMatserAPI:BaseUrl"];
                string AuthToken = _configuration["EWayBillMatserAPI:AuthToken"];
                string apiUrl = $"{baseUrl}{AuthToken}?{Parameters}";

                // API call
                var response = await httpClient.GetAsync(apiUrl);
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);

                // save API data to database
                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                var headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketnumber;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = Parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _sLEWayBillBusiness.saveEwayBillAPIsData(ApiData);

                // Validate response

                Console.WriteLine($"Response 1 : {responseData}");
                Console.WriteLine($"Result 1 : {result}");
                //Console.ReadKey();



                #endregion

                foreach (var date in dates)
                {
                    #region  2nd API call
                    // Parameters
                    string Parameters2 = $"email={Email}&date={date}";

                    // Headers
                    var httpClient2 = _httpClientFactory.CreateClient();
                    httpClient2.DefaultRequestHeaders.Add("client_id", clientId);
                    httpClient2.DefaultRequestHeaders.Add("client_secret", clientSecret);
                    httpClient2.DefaultRequestHeaders.Add("gstin", ClientGSTIN);
                    httpClient2.DefaultRequestHeaders.Add("ip_address", ipAddress);

                    // Body
                    RequestBody = new { };

                    // Url
                    string TogetEWayBill = _configuration["EWayBillMatserAPI:TogetEWayBill"];
                    apiUrl = $"{baseUrl}{TogetEWayBill}?{Parameters2}";

                    // API call
                    var response2 = await httpClient2.GetAsync(apiUrl);
                    var responseData2 = await response2.Content.ReadAsStringAsync();
                    var result2 = JsonConvert.DeserializeObject<JObject>(responseData2);

                    // save API data to database
                    var headersDict2 = httpClient2.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                    var headersJson2 = System.Text.Json.JsonSerializer.Serialize(headersDict2);
                    APIsDataModel ApiData2 = new APIsDataModel();
                    {
                        ApiData2.ClientGstin = ClientGSTIN;
                        ApiData2.RequestNumber = ticketnumber;
                        ApiData2.SessionID = sessionId;
                        ApiData2.RequestURL = apiUrl;
                        ApiData2.RequestParameters = Parameters2;
                        ApiData2.RequestHeaders = headersJson2;
                        ApiData2.RequestBody = JsonConvert.SerializeObject(RequestBody);
                        ApiData2.Response = responseData2;
                        ApiData2.ResponseCode = $"{(int)response2.StatusCode} {response2.StatusCode}";
                        ApiData2.Status = getStatus((int)response2.StatusCode);
                    }
                    ;
                    await _sLEWayBillBusiness.saveEwayBillAPIsData(ApiData2);

                    // Validate response
                    Console.WriteLine($"Response 2 : {responseData2}");
                    Console.WriteLine($"Result 2 : {result2}");
                    //Console.ReadKey();
                    List<string> EwbNos = new List<string>();
                    if (result2["status_cd"]?.ToString() == "1" && result2["data"] != null && result2["data"].HasValues)
                    {
                        EwbNos = GetEwbNos_Master(responseData2);
                    }
                    else
                    {
                        continue;
                    }

                    #endregion

                    foreach (var ewbno in EwbNos)
                    {
                        #region 3rd API call
                        // Parameters
                        string Parameters3 = $"email={Email}&ewbNo={ewbno}";

                        // Headers
                        var httpClient3 = _httpClientFactory.CreateClient();
                        httpClient3.DefaultRequestHeaders.Add("client_id", clientId);
                        httpClient3.DefaultRequestHeaders.Add("client_secret", clientSecret);
                        httpClient3.DefaultRequestHeaders.Add("gstin", ClientGSTIN);
                        httpClient3.DefaultRequestHeaders.Add("ip_address", ipAddress);
                        httpClient3.DefaultRequestHeaders.Add("Accept", accept);

                        // Body
                        RequestBody = new { };

                        // Url
                        string TogetEWayBIllDetails = _configuration["EWayBillMatserAPI:TogetEWayBIllDetails"];
                        apiUrl = $"{baseUrl}{TogetEWayBIllDetails}?{Parameters3}";

                        // API call
                        var response3 = await httpClient3.GetAsync(apiUrl);
                        var responseData3 = await response3.Content.ReadAsStringAsync();
                        var result3 = JsonConvert.DeserializeObject<JObject>(responseData3);

                        // save API data to database
                        var headersDict3 = httpClient3.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                        var headersJson3 = System.Text.Json.JsonSerializer.Serialize(headersDict3);
                        APIsDataModel ApiData3 = new APIsDataModel();
                        {
                            ApiData3.ClientGstin = ClientGSTIN;
                            ApiData3.RequestNumber = ticketnumber;
                            ApiData3.SessionID = sessionId;
                            ApiData3.RequestURL = apiUrl;
                            ApiData3.RequestParameters = Parameters3;
                            ApiData3.RequestHeaders = headersJson3;
                            ApiData3.RequestBody = JsonConvert.SerializeObject(RequestBody);
                            ApiData3.Response = responseData3;
                            ApiData3.ResponseCode = $"{(int)response3.StatusCode} {response3.StatusCode}";
                            ApiData3.Status = getStatus((int)response3.StatusCode);
                        }
                        ;
                        await _sLEWayBillBusiness.saveEwayBillAPIsData(ApiData3);

                        // Validate response
                        Console.WriteLine($"Response 3 : {responseData3}");
                        Console.WriteLine($"Result 3 : {result3}");
                        //Console.ReadKey();

                        if (result2["status_cd"]?.ToString() == "1" && result2["data"] != null && result2["data"].HasValues)
                        {
                            SL_Master_EWayBillJsontoDataTable(result3, SLEWayBillData);
                        }
                        else
                        {
                            continue;
                        }

                        #endregion
                    }
                }
                ClientGSTIN = tempClientGSTIN;
                await _sLEWayBillBusiness.SaveEWayBillDataAsync(SLEWayBillData, ticketnumber, ClientGSTIN);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"{ex.Message} ";
                return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTAPI_Master.cshtml");
            }
            #endregion

            // Comparison logic (unchanged)
            DataTable SLInvoice = await _sLDataBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            DataTable SLWayBill = await _sLEWayBillBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            DataTable SLEInvoice = await _sLEInvoiceBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);

            //Console.WriteLine("SLInvoice Columns: " + string.Join(", ", SLInvoice.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
            //Console.WriteLine("SLWayBill Columns: " + string.Join(", ", SLWayBill.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
            //Console.WriteLine("SLEInvoice Columns: " + string.Join(", ", SLEInvoice.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
            var EInvoice_AdminFileName = "EInvoice_master_API.csv";
            var EWayBill_AdminFileName = "EWayBill_master_API.csv";

            List<string> catagorys = new List<string>();
            var SLMatchTypes = _configuration["SLMatchTypes"].Split(',').Select(x => x.Trim()).ToList();

            await _sLComparedDataBusiness.CompareDataAsync(SLInvoice, SLWayBill, SLEInvoice);
            await _sLTicketsBusiness.UpdateSLTicketAsync(ticketnumber, ClientGSTIN, EInvoice_AdminFileName, EWayBill_AdminFileName);

            var data = await _sLComparedDataBusiness.GetComparedDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            ViewBag.ReportDataList = data;
            //_logger.LogInformation("Total Compared Data Rows: " + data.Count);

            // store number and sum in a model and save it in 
            ViewBag.Summary = GenerateSLSummary(data);
            ViewBag.GrandTotal = getGrandTotal(data);
            ViewBag.MatchType = SLMatchTypes;
            // var summary = ViewBag.Summary;

            return View("~/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedger.cshtml");
        }

        private DataTable SL_Master_EInvoiceJsontoDataTable(string json, string userGstin)
        {
            var table = new DataTable();

            // Define columns
            table.Columns.Add("User GSTIN");
            table.Columns.Add("GstRegType");
            table.Columns.Add("invoiceno");
            table.Columns.Add("invoice_date");
            table.Columns.Add("Supplier GSTIN");
            table.Columns.Add("SupplierName");
            table.Columns.Add("IsRcmApplied");
            table.Columns.Add("InvoiceValue", typeof(decimal));
            table.Columns.Add("ItemTaxableValue", typeof(decimal));
            table.Columns.Add("GstRate", typeof(decimal));
            table.Columns.Add("IGSTAmount", typeof(decimal));
            table.Columns.Add("CGSTAmount", typeof(decimal));
            table.Columns.Add("SGSTAmount", typeof(decimal));
            table.Columns.Add("CESS", typeof(decimal));
            table.Columns.Add("IsReturnFiled");
            table.Columns.Add("ReturnPeriod");

            JObject root = JObject.Parse(json);
            var retPeriod = root["header"]?["ret_period"]?.ToString();

            var b2bList = root["data"]?["b2b"] as JArray;
            if (b2bList == null) return table;

            foreach (var b2b in b2bList)
            {
                string cfs = b2b["cfs"]?.ToString(); // IsReturnFiled
                string ctin = b2b["ctin"]?.ToString();
                var invoices = b2b["inv"] as JArray;

                foreach (var inv in invoices)
                {
                    string inum = inv["inum"]?.ToString();
                    string idt = inv["idt"]?.ToString();
                    string inv_typ = inv["inv_typ"]?.ToString();
                    string rchrg = inv["rchrg"]?.ToString();
                    decimal val = inv["val"]?.ToObject<decimal>() ?? 0;

                    var items = inv["itms"] as JArray;
                    foreach (var item in items)
                    {
                        var details = item["itm_det"];
                        decimal txval = details["txval"]?.ToObject<decimal>() ?? 0;
                        decimal rt = details["rt"]?.ToObject<decimal>() ?? 0;
                        decimal iamt = details["iamt"]?.ToObject<decimal>() ?? 0;
                        decimal camt = details["camt"]?.ToObject<decimal>() ?? 0;
                        decimal samt = details["samt"]?.ToObject<decimal>() ?? 0;
                        decimal csamt = details["csamt"]?.ToObject<decimal>() ?? 0;

                        table.Rows.Add(
                           userGstin,         // User GSTIN
                           inv_typ,           // GstRegType
                           inum,              // invoiceno
                           idt,               // invoice_date
                           ctin,              // SupplierGSTIN
                           "",                // SupplierName (Not provided)
                           "",                // IsRcmApplied (Not provided)
                           val,               // InvoiceValue
                           txval,             // ItemTaxableValue
                           rt,                // GstRate
                           iamt,              // IGSTAmount
                           camt,              // CGSTAmount
                           samt,              // SGSTAmount
                           csamt,             // CESS
                           cfs,                // IsReturnFiled (Not provided)
                           retPeriod          // ReturnPeriod
                       );
                    }
                }
            }

            return table;
        }
        public static List<string> GetEwbNos_Master(string json)
        {
            var ewbNos = new List<string>();
            using var doc = JsonDocument.Parse(json);
            // Access the "data" array instead of "decrypted_data"
            var root = doc.RootElement.GetProperty("data");
            foreach (var item in root.EnumerateArray())
            {
                if (item.TryGetProperty("ewbNo", out var ewbNoProp))
                {
                    // Use GetInt64 since ewbNo is a number, then convert to string
                    ewbNos.Add(ewbNoProp.GetInt64().ToString());
                }
            }
            return ewbNos;
        }
        public void SL_Master_EWayBillJsontoDataTable(JObject result, DataTable EWBDataTable)
        {
            string[] allowedFormats = _configuration["Date Format"].Split(',').Select(s => s.Trim()).ToArray();
            var decryptedData = result["data"] as JObject;  // ✅ fixed

            if (decryptedData == null) return; // safeguard

            DataRow row = EWBDataTable.NewRow();

            row["EWB No"] = decryptedData["ewbNo"]?.ToString();

            // EWB Date
            string dateString = decryptedData["ewayBillDate"]?.ToString();
            if (DateTime.TryParseExact(dateString, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime ewbDate))
                row["EWB Date"] = ewbDate.ToString("dd-MM-yyyy hh:mm:ss tt");

            row["Supply Type"] = decryptedData["supplyType"]?.ToString();
            row["Doc.No"] = decryptedData["docNo"]?.ToString();

            // Doc Date
            string dateString2 = decryptedData["docDate"]?.ToString();
            if (DateTime.TryParseExact(dateString2, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime docDate))
                row["Doc.Date"] = docDate.ToString("dd-MM-yyyy hh:mm:ss tt");

            row["Doc.Type"] = decryptedData["docType"]?.ToString();
            row["TO GSTIN"] = decryptedData["toGstin"]?.ToString();
            row["status"] = decryptedData["status"]?.ToString();
            row["No of Items"] = decryptedData["itemList"]?.Count() ?? 0;

            // Get first item for HSN Code and Description
            var firstItem = decryptedData["itemList"]?.First;
            row["Main HSN Code"] = firstItem?["hsnCode"]?.ToString();
            row["Main HSN Desc"] = firstItem?["productDesc"]?.ToString();

            // Sum of taxable amounts
            double assessableValue = 0;
            foreach (var item in decryptedData["itemList"] ?? Enumerable.Empty<JToken>())
            {
                assessableValue += (double?)item["taxableAmount"] ?? 0;
            }
            row["Assessable Value"] = assessableValue;

            row["SGST Value"] = decryptedData["sgstValue"]?.ToString();
            row["CGST Value"] = decryptedData["cgstValue"]?.ToString();
            row["IGST Value"] = decryptedData["igstValue"]?.ToString();
            row["CESS Value"] = decryptedData["cessValue"]?.ToString();
            row["Total Invoice Value"] = decryptedData["totInvValue"]?.ToString();

            // Valid Till Date
            string dateString3 = decryptedData["validUpto"]?.ToString();
            if (DateTime.TryParseExact(dateString3, allowedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime validTillDate))
                row["Valid Till Date"] = validTillDate.ToString("dd-MM-yyyy hh:mm:ss tt");

            row["Gen.Mode"] = decryptedData["genMode"]?.ToString();

            EWBDataTable.Rows.Add(row);
        }

        #endregion

        #region SalesLedgerCurrentRequestsAPI-Iris 
        public async Task<IActionResult> SalesLedgerCurrentRequestsAPI_Iris(DateTime? fromdate, DateTime? todate)
        {
            DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
            DateTime toDateTime = todate ?? DateTime.Now; // Default to today

            //await TjCaptions("OpenTasks");
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                return View("~/Views/Admin/SalesLedgerCurrentRequests/SalesLedgerCurrentRequestsAPI_Iris.cshtml");
            }

            //var email = MySession.Current.Email; // Get the email from session
            //var clients = await _userBusiness.GetAdminClients(email);
            //string[] gstinArray = clients.Select(c => c.ClientGSTIN).ToArray();
            //var Tickets = await _sLTicketsBusiness.GetClientsOpenTicketsStatusAsync(gstinArray, fromDateTime, toDateTime);

            var Tickets = await _sLTicketsBusiness.GetOpenTicketsStatusAsync(MySession.Current.gstin, fromDateTime, toDateTime); //(ViewBag.UserGSTIN); // Fetch pending tickets based on user_gstin

            ViewBag.Tickets = Tickets;

            return View("~/Views/Admin/SalesLedgerCurrentRequests/SalesLedgerCurrentRequestsAPI_Iris.cshtml");
        }

        public async Task<IActionResult> CompareSalesLedgerGSTAPI_IRIS(string ticketnumber, string ClientGSTIN, string fromDate, string toDate)
        {
            // _logger.LogInformation($"Ticket Number: {ticketnumber}");
            //based on this ticket number fetch data from Purchase Ticket table and store
            var Ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);

            ViewBag.Ticket = Ticket;
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            ViewBag.fromDate = fromDate;
            ViewBag.toDate = toDate;

            return View("~/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTAPI_IRIS.cshtml");
        }

        [HttpPost]
        public async Task<IActionResult> CompareSalesLedgerAPI_IRIS(string ticketnumber, string ClientGSTIN)
        {
            string sessionId = HttpContext.Session.Id;
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act6";
            var Ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            ViewBag.Ticket = Ticket;

            string Txn_Period = Ticket.Period;
            string userName = "";
            try
            {
                var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);
                userName = UserAPIData.GstPortalUsername;
            }
            catch
            {
                string errorMessage = "Error fetching user API data. Please update user API data.";
                return Json(new { message = errorMessage });
            }

            string clientid = _configuration["EInvoiceIRISAPI:ClientId"];
            string clientSecret = _configuration["EInvoiceIRISAPI:ClientSecret"];
            string statecd = ClientGSTIN.Substring(0, 2);  // Extracting state code from GSTIN
            string ipusr = _configuration["EInvoiceIRISAPI:ipurs"];
            string txn = _configuration["EInvoiceIRISAPI:txn"];

            string parameters;
            var httpClient = _httpClientFactory.CreateClient();
            string baseUrl, AuthUrl, apiUrl, headersJson;

            try
            {
                // Step 1: Get existing auth token from DB

                // get authtoken and authtokenCreateddatetime and expiry from Database  - gstin,ticketno
                var authTokenData = await _gSTR2DataBusiness.GetTokenDataAsync(ClientGSTIN);

                bool isTokenExpired = authTokenData == null ||
                                      string.IsNullOrEmpty(authTokenData.AuthToken) ||
                                      (authTokenData.AuthTokenCreatedDatetime?.AddSeconds(Convert.ToDouble(authTokenData.Expiry)) < DateTime.Now);

                //bool isTokenExpired = authTokenData == null ||
                //					  string.IsNullOrEmpty(authTokenData.AuthToken) ||
                //					  (authTokenData.AuthTokenCreatedDatetime?.AddMinutes(Convert.ToDouble(authTokenData.Expiry)) < DateTime.Now);

                //Console.WriteLine($"ticketnumber : {ticketnumber}");
                //Console.WriteLine($"ClientGSTIN : {ClientGSTIN}");

                //Console.WriteLine($"AuthTokenCreatedDatetime: {authTokenData.AuthTokenCreatedDatetime}");
                //Console.WriteLine($"Expiry: {authTokenData.Expiry}");
                //Console.WriteLine($"Auth token valid upto : {authTokenData.AuthTokenCreatedDatetime?.AddSeconds(Convert.ToDouble(authTokenData.Expiry))}");
                //Console.WriteLine($"current time : {DateTime.Now}");

                if (!isTokenExpired)
                {
                    // Check buffer time (less than 1 hour?)
                    var bufferTime = (authTokenData.AuthTokenCreatedDatetime?.AddSeconds(Convert.ToDouble(authTokenData.Expiry)) - DateTime.Now);

                    //var bufferTime = (authTokenData.AuthTokenCreatedDatetime?.AddMinutes(Convert.ToDouble(authTokenData.Expiry)) - DateTime.Now);
                    //Console.WriteLine($"Buffer Time: {bufferTime}");
                    //Console.ReadKey();

                    if (bufferTime <= TimeSpan.FromHours(1))
                    {

                        #region 3rd API for refresh token

                        // Parametres
                        parameters = "";

                        // Headers
                        httpClient = _httpClientFactory.CreateClient();
                        httpClient.DefaultRequestHeaders.Add("clientid", clientid);
                        httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
                        httpClient.DefaultRequestHeaders.Add("state-cd", statecd);
                        httpClient.DefaultRequestHeaders.Add("ip-usr", ipusr);
                        httpClient.DefaultRequestHeaders.Add("txn", txn);

                        // Body
                        var RequestBody3 = new
                        {
                            action = "REFRESHTOKEN",
                            username = userName,
                            auth_token = authTokenData.AuthToken,
                            key = authTokenData.XAppKey,
                            sek = authTokenData.SEK
                        };
                        var content3 = new StringContent(JsonConvert.SerializeObject(RequestBody3), Encoding.UTF8, "application/json");

                        // Url
                        baseUrl = _configuration["EInvoiceIRISAPI:BaseUrl"];
                        AuthUrl = _configuration["EInvoiceIRISAPI:RefreshAuthToken"];
                        apiUrl = $"{baseUrl}{AuthUrl}";

                        // API call - Response
                        var response3 = await httpClient.PostAsync(apiUrl, content3);
                        var responseData3 = await response3.Content.ReadAsStringAsync();
                        var result3 = JsonConvert.DeserializeObject<JObject>(responseData3);

                        //Console.WriteLine($"Response 3 : {response3}");
                        //Console.WriteLine($"Result 3 : {result3}");
                        //Console.ReadKey();

                        // save API data to database
                        var headersDict3 = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                        headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict3);
                        APIsDataModel ApiData3 = new APIsDataModel();
                        {
                            ApiData3.ClientGstin = ClientGSTIN;
                            ApiData3.RequestNumber = ticketnumber;
                            ApiData3.SessionID = sessionId;
                            ApiData3.RequestURL = apiUrl;
                            ApiData3.RequestParameters = parameters;
                            ApiData3.RequestHeaders = headersJson;
                            ApiData3.RequestBody = JsonConvert.SerializeObject(RequestBody3);
                            ApiData3.Response = responseData3;
                            ApiData3.ResponseCode = $"{(int)response3.StatusCode} {response3.StatusCode}";
                            ApiData3.Status = getStatus((int)response3.StatusCode);
                        }
                        ;
                        await _sLEInvoiceBusiness.saveEInvoiceAPIsData(ApiData3);

                        // validate response
                        if (result3["status_cd"].ToString() != "1")
                        {
                            throw new Exception("Refresh Auth Token API Call - Failed to retrieve 'auth_token' from the response.");
                        }

                        string authToken = result3["auth_token"].ToString();
                        string expiry = result3["expiry"].ToString();
                        string sek = result3["sek"].ToString();

                        // Save token info  into DB
                        await _gSTR2DataBusiness.updateTokenData(new GSTR2TokenDataModel
                        {
                            ClientGstin = ClientGSTIN,
                            RequestNumber = ticketnumber,
                            UserName = userName,
                            XAppKey = authTokenData.XAppKey,
                            OTP = null,
                            AuthToken = authToken,
                            Expiry = expiry,
                            SEK = sek
                        });

                        #endregion

                    }

                    // Token is still valid, proceed to 4th API (later step)
                    //Console.WriteLine("Token is still valid. Proceeding to 4th API call.");
                    return Json(new { success = true, message = "Token is valid. Continue." });

                }

                #region Request for OTP         1st Api call 
                // Parameters
                parameters = "";

                // Headers
                httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Add("clientid", clientid);
                httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
                httpClient.DefaultRequestHeaders.Add("state-cd", statecd);
                httpClient.DefaultRequestHeaders.Add("ip-usr", ipusr);
                httpClient.DefaultRequestHeaders.Add("txn", txn);

                // Body
                var RequestBody1 = new
                {
                    action = "OTPREQUEST",
                    username = userName
                };
                var content = new StringContent(JsonConvert.SerializeObject(RequestBody1), Encoding.UTF8, "application/json");

                // Url
                baseUrl = _configuration["EInvoiceIRISAPI:BaseUrl"];
                AuthUrl = _configuration["EInvoiceIRISAPI:RequestForOTP"];
                apiUrl = $"{baseUrl}{AuthUrl}";

                // API call - Response
                var response = await httpClient.PostAsync(apiUrl, content);
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);
                //Console.WriteLine($"Response 1 : {response}");
                //Console.WriteLine($"Result 1 : {result}");
                //Console.ReadKey();

                // save API data to database   APIsDataModel
                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketnumber;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody1);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _sLEInvoiceBusiness.saveEInvoiceAPIsData(ApiData);
                //Console.WriteLine($"Client Gstin : {ClientGSTIN} ");
                //Console.WriteLine($"Request Number : {ticketnumber}");
                //Console.WriteLine($"Session ID : {sessionId}");
                //Console.WriteLine($"Request URL : {apiUrl}");
                //Console.WriteLine($"Request Parameters : {parameters}");
                //Console.WriteLine($"Request Headers: {headersJson}");
                //Console.WriteLine($"Request Body : {JsonConvert.SerializeObject(RequestBody1)}");
                //Console.WriteLine($"Response : {responseData}");
                //Console.WriteLine($"Response Code : {(int)response.StatusCode} {response.StatusCode}");
                //Console.WriteLine($"Status : {getStatus((int)response.StatusCode)}");

                //return Json(new
                //{
                //    success = false,
                //    askForOtp = true
                //});

                if (result["status_cd"].ToString() != "1")
                {
                    return Json(new
                    {
                        success = false,
                        askForOtp = false,
                        message = "OTP Request API Call - Failed to retrieve 'X-App-Key' from the response."
                    });
                    //throw new Exception("1st API Call - Failed to retrieve 'X-App-Key' from the response.");
                }

                //if (!xAppKeyResponse.IsSuccess)
                //    return Json(new { success = false, message = "Failed to get x-app-key from 1st API" });

                string x_app_key = result["x-app-key"].ToString();
                //Console.WriteLine($"x_app_key : {x_app_key}");

                #endregion

                // Return to front-end to prompt for OTP
                return Json(new
                {
                    success = false,
                    askForOtp = true,
                    xAppKey = x_app_key
                });
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"Error: {ex.Message}";
                return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTAPI_IRIS.cshtml");
            }

        }
        
		[HttpPost]
        public async Task<IActionResult> SLSubmitOtpAndContinue(string ClientGSTIN, string ticketNo, string otp, string xAppKey)
        {
            //Console.WriteLine("HI");
            string sessionId = HttpContext.Session.Id;
            ViewBag.Messages = "Admin";

            var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);

            try
            {
                #region Request For AuthToken            2nd api call
                string userName = UserAPIData.GstPortalUsername;
                string clientid = _configuration["EInvoiceIRISAPI:ClientId"];
                string clientSecret = _configuration["EInvoiceIRISAPI:ClientSecret"];
                string statecd = ClientGSTIN.Substring(0, 2);  // Extracting state code from GSTIN
                string ipusr = _configuration["EInvoiceIRISAPI:ipurs"];
                string txn = _configuration["EInvoiceIRISAPI:txn"];

                // Parameters
                string parameters = "";

                // Headers
                var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Add("clientid", clientid);
                httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
                httpClient.DefaultRequestHeaders.Add("state-cd", statecd);
                httpClient.DefaultRequestHeaders.Add("ip-usr", ipusr);
                httpClient.DefaultRequestHeaders.Add("txn", txn);

                // Body
                var RequestBody1 = new
                {
                    action = "AUTHTOKEN",
                    username = userName,
                    otp = otp,
                    key = xAppKey
                };
                var content = new StringContent(JsonConvert.SerializeObject(RequestBody1), Encoding.UTF8, "application/json");

                // Url
                string baseUrl = _configuration["EInvoiceIRISAPI:BaseUrl"];
                string AuthUrl = _configuration["EInvoiceIRISAPI:RequestForAuthToken"];
                string apiUrl = $"{baseUrl}{AuthUrl}";

                // API call - Response
                var response = await httpClient.PostAsync(apiUrl, content);
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);
                //Console.WriteLine($"Response 2 : {response}");
                //Console.WriteLine($"Result 2 : {result}");
                //Console.ReadKey();

                // save API data to database   APIsDataModel
                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                string headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketNo;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody1);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _sLEInvoiceBusiness.saveEInvoiceAPIsData(ApiData);

                //return Json(new
                //{
                //	success = false,
                //	askAgain = true,
                //	message = "Invalid OTP. Please enter again."
                //});

                // validate response
                if (result["status_cd"].ToString() != "1")
                {
                    if (result["error"]["error_cd"].ToString() == "AUTH4033")
                    {
                        return Json(new
                        {
                            success = false,
                            askAgain = true,
                            message = "Invalid OTP. Please enter again."
                        });
                    }
                    return Json(new { success = false, message = "AuthToken API Call - Failed to retrieve 'Token' from the response." });

                }

                string authToken = result["auth_token"].ToString();
                string expiry = result["expiry"].ToString();
                string sek = result["sek"].ToString();

                // Step: Save token info
                await _gSTR2DataBusiness.saveTokenData(new GSTR2TokenDataModel
                {
                    ClientGstin = ClientGSTIN,
                    RequestNumber = ticketNo,
                    UserName = userName,
                    XAppKey = xAppKey,
                    OTP = otp,
                    AuthToken = authToken,
                    Expiry = expiry,
                    SEK = sek
                });
                #endregion

                return Json(new { success = true, message = "OTP Verified. Token saved." });

            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"Error: {ex.Message}";
                return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTAPI_IRIS.cshtml");
            }

        }

        public async Task<IActionResult> SLContinueWith4thApi(string ticketnumber, string clientGSTIN)
        {
            //Console.WriteLine($"TicketNumber  4: {ticketnumber}");
            //Console.WriteLine($"ClientGSTIN 4: {clientGSTIN}");

            string sessionId = HttpContext.Session.Id;
            ViewBag.Messages = "Admin";
            var Ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, clientGSTIN);
            ViewBag.Ticket = Ticket;

            var authTokenData = await _gSTR2DataBusiness.GetTokenDataAsync(clientGSTIN);

            DataTable EInvoiceDataTable = new DataTable();
            // Read headers
            var EInvoiceColumns = _configuration["SLEInvoice"].Split(',').Select(x => x.Trim()).ToList();
            foreach (string header in EInvoiceColumns)
            {
                EInvoiceDataTable.Columns.Add(header.Trim(), typeof(string));
            }

            string txnPeriod = Ticket.Period;
            string formattedPeriod = DateTime.ParseExact(txnPeriod, "MMM-yy", CultureInfo.InvariantCulture).ToString("MMyyyy");


            try
            {
                #region 4th api call 
                string clientid = _configuration["EInvoiceIRISAPI:ClientId"];
                string clientSecret = _configuration["EInvoiceIRISAPI:ClientSecret"];
                string statecd = clientGSTIN.Substring(0, 2);  // Extracting state code from GSTIN
                string ipusr = _configuration["EInvoiceIRISAPI:ipurs"];
                string txn = _configuration["EInvoiceIRISAPI:txn"];

                // Parameters
                string action = "IRNJSON";
                string rtnprd = formattedPeriod;
                string suptyp = "B2B";
                string rtin = clientGSTIN;

                //action: IRNJSON
                //rtnprd:042025
                //suptyp: B2B
                //rtin:33AQIPK2639R1ZB

                string parameter = $"action={action}&rtnprd={rtnprd}&suptyp={suptyp}&rtin={rtin}";

                // Headers
                var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Add("clientid", clientid);
                httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
                httpClient.DefaultRequestHeaders.Add("auth-token", authTokenData.AuthToken);
                httpClient.DefaultRequestHeaders.Add("username", authTokenData.UserName);
                httpClient.DefaultRequestHeaders.Add("state-cd", statecd);
                httpClient.DefaultRequestHeaders.Add("ip-usr", ipusr);
                httpClient.DefaultRequestHeaders.Add("txn", txn);
                httpClient.DefaultRequestHeaders.Add("gstin", authTokenData.ClientGstin);
                httpClient.DefaultRequestHeaders.Add("ret-period", formattedPeriod);
                httpClient.DefaultRequestHeaders.Add("x-sek", authTokenData.SEK);
                httpClient.DefaultRequestHeaders.Add("x-app-key", authTokenData.XAppKey);

                //clientid: IRIS12a90c8d1e17487e65e1b31a490d398e
                //client-secret:4538bfce2d10eeddee3b0f07596cf65b
                //auth-token:63f2e50713a44257b6a90c8d35bda95c
                //username:ragulindustrie
                //state - cd:33
                //ip - usr:27.168.1.1
                //txn: AKSPY7021609661
                //gstin:33AQIPK2639R1ZB
                //ret - period:042025
                //Content - Type:application / json
                //x - sek:3xWvKJN6C0s / HJ4r3ATBjoYLfqlmM3z95aeFpBw6M1hr / yKC0jFhgzLlWove9k0X
                //x - app - key:SQ3YjtLQ4wbhNH3KSI13K + Q3IapBMJwiVGHJI3Yeg6A =

                // Body
                var RequestBody1 = new
                {

                };
                var content = new StringContent(JsonConvert.SerializeObject(RequestBody1), Encoding.UTF8, "application/json");

                // Url
                string baseUrl = _configuration["EInvoiceIRISAPI:BaseUrl"];
                string AuthUrl = _configuration["EInvoiceIRISAPI:GetIRNJSON"];
                string apiUrl = $"{baseUrl}{AuthUrl}?{parameter}";

                // Api call - Response
                var response = await httpClient.GetAsync(apiUrl);
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);

                // save API data to database
                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                string headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel
                {
                    ClientGstin = clientGSTIN,
                    RequestNumber = ticketnumber,
                    SessionID = sessionId,
                    RequestURL = apiUrl,
                    RequestParameters = parameter,
                    RequestHeaders = headersJson,
                    RequestBody = JsonConvert.SerializeObject(RequestBody1),
                    Response = responseData,
                    ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}",
                    Status = getStatus((int)response.StatusCode)
                };
                await _sLEInvoiceBusiness.saveEInvoiceAPIsData(ApiData);

                Console.WriteLine($"Response 4 : {response}");
                Console.WriteLine($"Result 4 : {result}");

                // validate response

                //{
                //	"est": "30",
                //	"token": "2ffea879af264c36915fc7d5ba381a36"
                //}
                if (result["token"] != null)
                {
                    string token = result["token"]?.ToString();

                    return Json(new { success = true, token = token });
                    //return RedirectToAction("ContinueWith5thApi", new { ticketnumber, clientGSTIN, token5 });

                }
                //{
                //	"status_code": 0,
                //	"message": "The file for the request date and type is already generated on 13/08/2025 11:37:17. To download the same file, use token 6001f2a9df354b2ca15532fc86c1d4ae. The link is valid till 1 day."
                //}
                else if (result["message"] != null)
                {
                    string message = result["message"]?.ToString();
					if (message.Contains("already generated", StringComparison.OrdinalIgnoreCase))
					{
						// Match any non-space sequence after the word "token"
						var match = System.Text.RegularExpressions.Regex.Match(message, @"token\s+([^\s.]+)", RegexOptions.IgnoreCase);
						if (match.Success)
						{
							string tokenFromMessage = match.Groups[1].Value;
							return Json(new { success = true, token = tokenFromMessage });
						}
					}
                }
                //{ "status_code": 0,
				//	"message": "Invalid Auth token or username"
				//}

                // {
                //   "status_code": 0,
                //   "message": "Invalid API Key"
                // }
                if (result["status_code"] != null && result["status_code"].ToString() == "0")
                {
                    string msg = result["message"]?.ToString();
                    return Json(new
                    {
                        failure = true,
                        message = $"\"Get IRN JSON\" API Call Failed - Response - \"{msg}\""
                    });

                }

                // convert response data to DataTable
                //EInvoiceDataTable = ConvertJsontoPRDataTable(responseData, clientGSTIN, EInvoiceDataTable);

                #endregion

                // Save the data to database
                //await _gSTR2DataBusiness.SaveGSTR2DataAsync(EInvoiceDataTable, ticketnumber, clientGSTIN);

            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"Error: {ex.Message}";
                return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTAPI_IRIS.cshtml");
            }

            return Json(new { goToCompare = true });
        }

        public async Task<IActionResult> ContinueWithFinalApi(string ticketnumber, string clientGSTIN, string token5)
        {
            ViewBag.Messages = "Admin";
            string sessionId = HttpContext.Session.Id;
            var ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, clientGSTIN);
            ViewBag.Ticket = ticket;

            var authTokenData = await _gSTR2DataBusiness.GetTokenDataAsync(clientGSTIN);

            DataTable EInvoiceDataTable = new DataTable();
            // Read headers
            var EInvoiceColumns = _configuration["SLEInvoice"].Split(',').Select(x => x.Trim()).ToList();
            foreach (string header in EInvoiceColumns)
            {
                EInvoiceDataTable.Columns.Add(header.Trim(), typeof(string));
            }

            string txnPeriod = ticket.Period;
            string formattedPeriod = DateTime.ParseExact(txnPeriod, "MMM-yy", CultureInfo.InvariantCulture).ToString("MMyyyy");


            #region Final API Call  for Einvoice
            try
            {

                #region  5th API call to get data
                string clientid = _configuration["EInvoiceIRISAPI:ClientId"];
                string clientSecret = _configuration["EInvoiceIRISAPI:ClientSecret"];
                string statecd = clientGSTIN.Substring(0, 2);  // Extracting state code from GSTIN
                string ipusr = _configuration["EInvoiceIRISAPI:ipurs"];
                string txn = _configuration["EInvoiceIRISAPI:txn"];

                // Parameters
                string action5 = "FILEDETL";
                string gstin5 = clientGSTIN;
                string ret_period = formattedPeriod;
                string parameter5 = $"action={action5}&gstin={gstin5}&ret_period={ret_period}&token={token5}";
                //action: FILEDETL
                //gstin:33AQIPK2639R1ZB
                //ret_period:042025
                //token: 80ff816e162c4f4e8bb5e905b2307970

                // Headers
                var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Add("clientid", clientid);
                httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
                httpClient.DefaultRequestHeaders.Add("auth-token", authTokenData.AuthToken);
                httpClient.DefaultRequestHeaders.Add("username", authTokenData.UserName);
                httpClient.DefaultRequestHeaders.Add("state-cd", statecd);
                httpClient.DefaultRequestHeaders.Add("ip-usr", ipusr);
                httpClient.DefaultRequestHeaders.Add("txn", txn);
                httpClient.DefaultRequestHeaders.Add("gstin", authTokenData.ClientGstin);
                httpClient.DefaultRequestHeaders.Add("ret-period", formattedPeriod);
                httpClient.DefaultRequestHeaders.Add("x-sek", authTokenData.SEK);
                httpClient.DefaultRequestHeaders.Add("x-app-key", authTokenData.XAppKey);

                //clientid: IRIS12a90c8d1e17487e65e1b31a490d398e
                //client - secret:4538bfce2d10eeddee3b0f07596cf65b
                //auth - token:4df9b2ba06c54cb181b174e9cfae64cc
                //username:ragulindustrie
                //state - cd:33
                //ip - usr:27.168.1.1
                //txn: AKSPY7021609661
                //gstin:33AQIPK2639R1ZB
                //ret - period:042025
                //Content - Type:application / json
                //x - sek:VF4OO5EL5Rmil9shTEq4A1vTKMOEAxKALmvC3nHt / aktOTW0UMU8QhL6v9IEKuc1
                //x - app - key:p4skJiAHaQltmOd76JJzEzg9FLKEtDJAjaERBxAtHaU =

                // Body
                var RequestBody5 = new
                {

                };
                var content5 = new StringContent(JsonConvert.SerializeObject(RequestBody5), Encoding.UTF8, "application/json");

                // Url
                string baseUrl5 = _configuration["EInvoiceIRISAPI:BaseUrl"];
                string AuthUrl5 = _configuration["EInvoiceIRISAPI:GetEinvoiceData"];
                string apiUrl5 = $"{baseUrl5}{AuthUrl5}?{parameter5}";

                // Api call - Response
                var response5 = await httpClient.GetAsync(apiUrl5);
                var responseData5 = await response5.Content.ReadAsStringAsync();
                var result5 = JsonConvert.DeserializeObject<JObject>(responseData5);

                // save API data to database
                var headersDict5 = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                string headersJson5 = System.Text.Json.JsonSerializer.Serialize(headersDict5);
                APIsDataModel ApiData5 = new APIsDataModel
                {
                    ClientGstin = clientGSTIN,
                    RequestNumber = ticketnumber,
                    SessionID = sessionId,
                    RequestURL = apiUrl5,
                    RequestParameters = parameter5,
                    RequestHeaders = headersJson5,
                    RequestBody = JsonConvert.SerializeObject(RequestBody5),
                    Response = responseData5,
                    ResponseCode = $"{(int)response5.StatusCode} {response5.StatusCode}",
                    Status = getStatus((int)response5.StatusCode)
                };
                await _sLEInvoiceBusiness.saveEInvoiceAPIsData(ApiData5);

                Console.WriteLine($"Response 5 : {response5}");
                Console.WriteLine($"Result 5 : {result5}");
                //Console.WriteLine($"1 - {result5["status_code"]?.ToString() == "0"} , 2 - {result5["urls"] != null}");

                //Result 5 : {
                //			  "ek": "yNRzdodhcNN9Alqmsr9CcQ5S8QOw4yTWcpksee9sQHc=",
                //            "urls": [
                //		            {
                //					  "ic": 1,
                //                    "ul": "https://files.gst.gov.in/einvdownloads/05022025/JSON/2c5847edad8642e4b42a794142fc63fd/27AAFCI5032C1ZZ_122024_Received_1.tar.gz?md5=Sa5nyD8SE6q5l2EMZmNmRg&expires=1738915665",
                //                    "hash": "6c1be18b84bff4ae0c61c719de7c5c1a9a6e141888cb8bd140de37bcb9382c6d"
                //		            }
                //                    ],
                //           "fc": 1,
                //           "nextAvailable": "2025-02-06 13:37:12"
                //          }
                if (result5["urls"] != null)
                {
                    string ek = result5["ek"]?.ToString();
                    int fileCount = (int)result5["fc"];
                    JArray URLs = (JArray)result5["urls"];
                    string[] urls = URLs.Select(u => u["ul"]?.ToString()).ToArray();

                    foreach (string url in urls)
                    {
                        //Console.WriteLine($"URL: {url}");
                        // Step 1: Download tar.gz file
                        using HttpClient client = new HttpClient();
                        byte[] fileBytes = await client.GetByteArrayAsync(url);

                        // Step 2: Save it temporarily
                        string tempTarGzPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".tar.gz");
                        await System.IO.File.WriteAllBytesAsync(tempTarGzPath, fileBytes);

                        // Step 3: Extract .tar.gz
                        string extractedFolderPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                        Directory.CreateDirectory(extractedFolderPath);

                        // First extract .gz to .tar
                        string tempTarPath = Path.ChangeExtension(tempTarGzPath, ".tar");
                        using (FileStream originalFileStream = new FileStream(tempTarGzPath, FileMode.Open, FileAccess.Read))
                        using (FileStream decompressedFileStream = new FileStream(tempTarPath, FileMode.Create))
                        using (GZipStream decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress))
                        {
                            decompressionStream.CopyTo(decompressedFileStream);
                        }

                        // Now extract .tar file
                        using (var archive = TarArchive.Open(tempTarPath))
                        {
                            foreach (var entry in archive.Entries.Where(entry => !entry.IsDirectory))
                            {
                                string extractedFilePath = Path.Combine(extractedFolderPath, entry.Key);
                                entry.WriteToFile(extractedFilePath);

                                // Step 4: Read JSON file content
                                string jsonContent = await System.IO.File.ReadAllTextAsync(extractedFilePath);

                                // Step 5: Pass to DecryptEinvoiceResponse
                                EInvoiceDataTable = DecryptEinvoiceResponse(jsonContent, ek, clientGSTIN, EInvoiceDataTable);
                            }
                        }

                        // Optional: Clean up temp files
                        System.IO.File.Delete(tempTarGzPath);
                        System.IO.File.Delete(tempTarPath);
                        Directory.Delete(extractedFolderPath, true);

                        //Console.WriteLine("Hi");

                    }

                    // Save the data to database
                    await _sLEInvoiceBusiness.SaveEInvoiceDataAsync(EInvoiceDataTable, ticketnumber, clientGSTIN);

                    return Json(new { success = true, ticketno = ticketnumber, gstin = clientGSTIN });

                }

                //Result 5 : {
                //           "status_code": 0,
                //           "message": "File generation is in progress"
                //           }
                if (result5["status_code"]?.ToString() == "0")
                {
                    string message = result5["message"]?.ToString();

                    if (!string.IsNullOrEmpty(message) && message.Contains("File generation"))
                    {

                        return Json(new { message = "File Generation", token = token5 });
                        //return View("/Views/Admin/CompareGstFiles/CompareGSTAPI_IRIS.cshtml");
                    }
                    else
                    {
                        return Json(new
                        {
                            error = true,
                            message = $"\"Get File link FILEDETL\" API Call Failed - Response - \"{message}\"",
                        });
                        //throw new Exception($"5th API Call - Failed due to {message}");
                    }
                }

                else if (result5["error_code"]?.ToString() == "500")
                {
                    string message = result5["message"]?.ToString();
                    message += result5["description"]?.ToString();
                    return Json(new
                    {
                        error = true,
                        message = message,
                    });
                }
                #endregion

            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"{ex.Message} ";
                return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTAPI_IRIS.cshtml");
            }
            #endregion

            return RedirectToAction("CompareSLDataTables", new { ticketnumber, clientGSTIN });
        }

        public async Task<IActionResult> CompareSLDataTables(string ticketnumber, string ClientGSTIN)
        {
            ViewBag.Messages = "Admin";
            string sessionId = HttpContext.Session.Id;
            var ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            ViewBag.Ticket = ticket;

            var period = ticket.Period;
            List<string> dates = new List<string>();
            string formate = "dd/MM/yyyy";
            // Parse the month and year
            DateTime startDate = DateTime.ParseExact("01-" + period, "dd-MMM-yy", null); //DateTime.ParseExact("01-Apr-25", "dd-MMM-yy", null) ;// 
            int daysInMonth = DateTime.DaysInMonth(startDate.Year, startDate.Month);
            // Loop through each day in the month
            for (int day = 1; day <= daysInMonth; day++)
            {
                DateTime currentDate = new DateTime(startDate.Year, startDate.Month, day);
                string formattedDate = currentDate.ToString(formate, CultureInfo.InvariantCulture);

                //string formattedDate = currentDate.ToString(formate);
                dates.Add(formattedDate);
            }
            List<string> otherdays = await _sLDataBusiness.GetOtherDates(ticketnumber, ClientGSTIN, period, formate);
            dates.AddRange(otherdays);
            //foreach (var date in otherdays)
            //{
            //    Console.WriteLine($"Date: {date}");
            //}    
            DataTable EWBDataTable = new DataTable();
            // Read headers
            var EWayBillColumns = _configuration["SLEWayBill"].Split(',').Select(x => x.Trim()).ToList();
            foreach (string header in EWayBillColumns)
            {
                EWBDataTable.Columns.Add(header.Trim(), typeof(string));
            }

            string apiUserName = "";
            string apiPassword = "";
            try
            {
                var UserAPIData = await _userBusiness.GetClientAPIData(ClientGSTIN);
                apiUserName = UserAPIData.APIIRISUsername;
                apiPassword = UserAPIData.APIIRISPassword;
            }
            catch (Exception e)
            {
                ViewBag.ErrorMessage = "Error fetching user API data. Please update user API data.";
                return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTAPI_IRIS.cshtml");

            }

            #region API call for EWayBill Data

            try
            {
                #region App key generator                1st API call for EWay Bill Data
                // parameters
                string parameters = "";
                // headers
                string clientid = _configuration["EWayBillIRISAPI:ClientId"];
                string clientSecret = _configuration["EWayBillIRISAPI:ClientSecret"];
                // body 			
                string username = apiUserName;  //"API_RAGULIRIS"; // get this data from db
                string password = apiPassword;  //"Enter@2025";    // get this data from db				
                var RequestBody1 = new
                {
                    type = "Public",
                    data = new
                    {
                        action = "AUTH",
                        username = username,
                        password = password
                    },
                    portalType = "ewaybill"
                };
                var content = new StringContent(JsonConvert.SerializeObject(RequestBody1), Encoding.UTF8, "application/json");

                string baseUrl = _configuration["EWayBillIRISAPI:BaseUrl"];
                string AppKeyGenerator = _configuration["EWayBillIRISAPI:AppKeyGenerator"];
                // Url
                string apiUrl = $"{baseUrl}{AppKeyGenerator}";

                var httpClient = _httpClientFactory.CreateClient();
                // Add required headers
                httpClient.DefaultRequestHeaders.Add("clientid", clientid);
                httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);


                var response = await httpClient.PostAsync(apiUrl, content);
                var responseData = await response.Content.ReadAsStringAsync();
                var result = JsonConvert.DeserializeObject<JObject>(responseData);

                //Console.WriteLine($"Response-1 : {result}");
                var headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                string headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                APIsDataModel ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketnumber;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody1);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _sLEWayBillBusiness.saveEwayBillAPIsData(ApiData);
                //Console.WriteLine($"Client Gstin : {ClientGSTIN} ");
                //Console.WriteLine($"Request Number : {ticketnumber}");
                //Console.WriteLine($"Session ID : {sessionId}");
                //Console.WriteLine($"Request URL : {apiUrl}");
                //Console.WriteLine($"Request Parameters : {parameters}");
                //Console.WriteLine($"Request Headers: {headersJson}");
                //Console.WriteLine($"Request Body : {JsonConvert.SerializeObject(RequestBody1)}");
                //Console.WriteLine($"Response : {responseData}");
                //Console.WriteLine($"Response Code : {(int)response.StatusCode} {response.StatusCode}");
                //Console.WriteLine($"Status : {getStatus((int)response.StatusCode)}");
                if (!result.ContainsKey("app_key") || string.IsNullOrEmpty(result["app_key"].ToString()))
                {
                    throw new Exception("1st API Call - Failed to retrieve 'app key' from the response.");
                }

                string appKey = result["app_key"].ToString();  // save in db

                #endregion

                #region  Encrypt Auth Payload      2nd API call for EWay Bill Data
                // parameters
                parameters = "";
                // Body
                var RequestBody2 = new
                {
                    type = "Public",
                    data = new
                    {
                        action = "ACCESSTOKEN",
                        username = username,
                        password = password,
                        app_key = appKey // This must be the Base64 string returned from the first call
                    },
                    portalType = "ewaybill"
                };
                content = new StringContent(JsonConvert.SerializeObject(RequestBody2), Encoding.UTF8, "application/json");

                string EncryptAuthPayload = _configuration["EWayBillIRISAPI:EncryptAuthPayload"];
                // Url
                apiUrl = $"{baseUrl}{EncryptAuthPayload}";

                response = await httpClient.PostAsync(apiUrl, content);
                responseData = await response.Content.ReadAsStringAsync();

                result = JsonConvert.DeserializeObject<JObject>(responseData);
                headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketnumber;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody2);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _sLEWayBillBusiness.saveEwayBillAPIsData(ApiData);
                if (!result.ContainsKey("Data") || string.IsNullOrEmpty(result["Data"]?.ToString()))
                {
                    throw new Exception("2nd API Call - Failed to retrieve 'Data' from the response.");
                }
                string Data = result["Data"]?.ToString();
                //Console.WriteLine($"Response-2 : {result}");

                #endregion

                #region Authorization For EWayBill        3rd API call for EWay Bill Data
                // parameters
                parameters = "";
                //header
                string gstin = "33AQIPK2639R1ZB"; // ClientGSTIN
                                                  // Body								  
                var RequestBody3 = new
                {
                    Data = Data
                };
                content = new StringContent(JsonConvert.SerializeObject(RequestBody3), Encoding.UTF8, "application/json");
                string AuthorizationForEWayBill = _configuration["EWayBillIRISAPI:AuthorizationForEWayBill"];
                //url
                apiUrl = $"{baseUrl}{AuthorizationForEWayBill}";
                // Add required headers
                httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Add("client-id", clientid);
                httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
                httpClient.DefaultRequestHeaders.Add("Gstin", gstin);

                response = await httpClient.PostAsync(apiUrl, content);
                responseData = await response.Content.ReadAsStringAsync();

                result = JsonConvert.DeserializeObject<JObject>(responseData);
                headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketnumber;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody3);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _sLEWayBillBusiness.saveEwayBillAPIsData(ApiData);

                //Console.WriteLine($"Response-3 : {result}");
                string authtoken, sek;
                if (result.ContainsKey("status") && result["status"].ToString() == "1")
                {
                    authtoken = result["authtoken"]?.ToString();
                    sek = result["sek"]?.ToString();
                }
                else
                {
                    throw new Exception("3rd API Call - Failed to retrieve 'authtoken' and 'sek' from the response.");
                }

                #endregion

                #region Decrypt SEK key       4th API call for EWay Bill Data
                // parameters
                parameters = "";
                //Body
                var RequestBody4 = new
                {
                    type = "Public",
                    key = appKey,
                    data = sek,
                    portalType = "ewaybill"
                };
                content = new StringContent(JsonConvert.SerializeObject(RequestBody4), Encoding.UTF8, "application/json");
                string DecryptSEKkey = _configuration["EWayBillIRISAPI:DecryptSEKkey"];
                // Url
                apiUrl = $"{baseUrl}{DecryptSEKkey}";
                // Add required headers
                httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Add("clientid", clientid);
                httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);

                response = await httpClient.PostAsync(apiUrl, content);
                responseData = await response.Content.ReadAsStringAsync();
                result = JsonConvert.DeserializeObject<JObject>(responseData);
                //Console.WriteLine($"Response-4 : {result}");
                headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                ApiData = new APIsDataModel();
                {
                    ApiData.ClientGstin = ClientGSTIN;
                    ApiData.RequestNumber = ticketnumber;
                    ApiData.SessionID = sessionId;
                    ApiData.RequestURL = apiUrl;
                    ApiData.RequestParameters = parameters;
                    ApiData.RequestHeaders = headersJson;
                    ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody4);
                    ApiData.Response = responseData;
                    ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                    ApiData.Status = getStatus((int)response.StatusCode);
                }
                ;
                await _sLEWayBillBusiness.saveEwayBillAPIsData(ApiData);
                string decryptkey;
                if (!result.ContainsKey("decrypt_sek_key") || string.IsNullOrEmpty(result["decrypt_sek_key"]?.ToString()))
                {
                    throw new Exception("4nd API Call - Failed to retrieve 'decrypt_sek_key' from the response.");
                }
                decryptkey = result["decrypt_sek_key"]?.ToString(); // save in db



                #endregion

                //int i = 1;
                foreach (var date in dates)
                {
                    #region GETEWayBillsByDate 	  5th API call for EWay Bill Data

                    //parameters
                    parameters = $"date={date}";
                    // Body
                    var RequestBody5 = new
                    {

                    };
                    // Url
                    string GETEWayBillsByDate = _configuration["EWayBillIRISAPI:GETEWayBillsByDate"];
                    apiUrl = $"{baseUrl}{GETEWayBillsByDate}?{parameters}";
                    // Add required headers
                    httpClient = _httpClientFactory.CreateClient();
                    httpClient.DefaultRequestHeaders.Add("client-id", clientid);
                    httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
                    httpClient.DefaultRequestHeaders.Add("gstin", gstin);
                    httpClient.DefaultRequestHeaders.Add("authtoken", authtoken);

                    response = await httpClient.GetAsync(apiUrl);
                    responseData = await response.Content.ReadAsStringAsync();
                    result = JsonConvert.DeserializeObject<JObject>(responseData);
                    //Console.WriteLine($"Response-5 : {result}");

                    headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                    headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                    ApiData = new APIsDataModel();
                    {
                        ApiData.ClientGstin = ClientGSTIN;
                        ApiData.RequestNumber = ticketnumber;
                        ApiData.SessionID = sessionId;
                        ApiData.RequestURL = apiUrl;
                        ApiData.RequestParameters = parameters;
                        ApiData.RequestHeaders = headersJson;
                        ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody5);
                        ApiData.Response = responseData;
                        ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                        ApiData.Status = getStatus((int)response.StatusCode);
                    }
                    ;
                    await _sLEWayBillBusiness.saveEwayBillAPIsData(ApiData);
                    string encryptedData, rek, hmac;
                    if (result.ContainsKey("status") && result["status"].ToString() == "1")
                    {
                        encryptedData = result["data"]?.ToString();
                        rek = result["rek"]?.ToString();
                        hmac = result["hmac"]?.ToString();
                    }
                    else
                    {
                        continue; // Skip to the next date if the API call fails
                        throw new Exception("5nd API Call - Failed to retrieve 'Encrypted Data','rek','hmac' from the response.");
                    }
                    #endregion

                    #region DecryptEWayBillsByDate  6th API call for EWay Bill Data
                    // parameters
                    parameters = "";
                    // Body
                    var RequestBody6 = new
                    {
                        type = "Private",
                        key = decryptkey,
                        data = encryptedData,
                        rek = rek,
                        portalType = "ewaybill"
                    };
                    content = new StringContent(JsonConvert.SerializeObject(RequestBody6), Encoding.UTF8, "application/json");
                    // Url
                    string DecryptEWayBillsByDate = _configuration["EWayBillIRISAPI:DecryptEWayBillsByDate"];
                    apiUrl = $"{baseUrl}{DecryptEWayBillsByDate}";
                    // Add required headers
                    httpClient = _httpClientFactory.CreateClient();
                    httpClient.DefaultRequestHeaders.Add("clientid", clientid);
                    httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);

                    response = await httpClient.PostAsync(apiUrl, content);
                    responseData = await response.Content.ReadAsStringAsync();
                    result = JsonConvert.DeserializeObject<JObject>(responseData);
                    //Console.WriteLine($"Response-6 : {result}");
                    headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                    headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                    ApiData = new APIsDataModel();
                    {
                        ApiData.ClientGstin = ClientGSTIN;
                        ApiData.RequestNumber = ticketnumber;
                        ApiData.SessionID = sessionId;
                        ApiData.RequestURL = apiUrl;
                        ApiData.RequestParameters = parameters;
                        ApiData.RequestHeaders = headersJson;
                        ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody6);
                        ApiData.Response = responseData;
                        ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                        ApiData.Status = getStatus((int)response.StatusCode);
                    }
                    ;
                    await _sLEWayBillBusiness.saveEwayBillAPIsData(ApiData);

                    if (!result.ContainsKey("decrypted_data") || string.IsNullOrEmpty(result["decrypted_data"]?.ToString()))
                    {
                        throw new Exception("6th API Call - Failed to retrieve 'decrypted_data' from the response.");
                    }


                    List<string> EwbNo = GetEwbNos(responseData);
                    foreach (var ewbno in EwbNo)
                    {
                        //i++;
                        //Console.WriteLine($"EWB No: {ewbno}");

                        #region GETEWayBillDetails          7th API call for EWay Bill Data

                        // parameters
                        parameters = $"ewbNo={ewbno}";
                        // Body
                        var RequestBody7 = new
                        {

                        };
                        // headers
                        httpClient = _httpClientFactory.CreateClient();
                        httpClient.DefaultRequestHeaders.Add("client-id", clientid);
                        httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
                        httpClient.DefaultRequestHeaders.Add("gstin", gstin);
                        httpClient.DefaultRequestHeaders.Add("authtoken", authtoken);
                        // Url
                        string GETEWayBillDetails = _configuration["EWayBillIRISAPI:GETEWayBillDetails"];
                        apiUrl = $"{baseUrl}{GETEWayBillDetails}?{parameters}";

                        response = await httpClient.GetAsync(apiUrl);
                        responseData = await response.Content.ReadAsStringAsync();
                        result = JsonConvert.DeserializeObject<JObject>(responseData);
                        //Console.WriteLine($"Response-7 : {result}");
                        headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                        headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                        ApiData = new APIsDataModel();
                        {
                            ApiData.ClientGstin = ClientGSTIN;
                            ApiData.RequestNumber = ticketnumber;
                            ApiData.SessionID = sessionId;
                            ApiData.RequestURL = apiUrl;
                            ApiData.RequestParameters = parameters;
                            ApiData.RequestHeaders = headersJson;
                            ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody7);
                            ApiData.Response = responseData;
                            ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                            ApiData.Status = getStatus((int)response.StatusCode);
                        }
                        ;
                        await _sLEWayBillBusiness.saveEwayBillAPIsData(ApiData);
                        string details_encryptedData, details_rek, details_hmac;
                        if (result.ContainsKey("status") && result["status"].ToString() == "1")
                        {
                            details_encryptedData = result["data"]?.ToString();
                            details_rek = result["rek"]?.ToString();
                            details_hmac = result["hmac"]?.ToString();
                        }
                        else
                        {
                            continue; // Skip to the next EWB number if the API call fails
                            throw new Exception("7th API Call - Failed to retrieve Encrypted EWay Bill details from the response.");
                        }
                        #endregion

                        #region  DecryptEWayBillDetails        8th API call for EWay Bill Data
                        // parameters
                        parameters = "";
                        // Add required headers
                        httpClient = _httpClientFactory.CreateClient();
                        httpClient.DefaultRequestHeaders.Add("clientid", clientid);
                        httpClient.DefaultRequestHeaders.Add("client-secret", clientSecret);
                        //body 
                        var RequestBody8 = new
                        {
                            type = "Private",
                            key = decryptkey,
                            data = details_encryptedData,
                            rek = details_rek,
                            portalType = "ewaybill"
                        };
                        content = new StringContent(JsonConvert.SerializeObject(RequestBody8), Encoding.UTF8, "application/json");
                        // url
                        string DecryptEWayBillDetails = _configuration["EWayBillIRISAPI:DecryptEWayBillDetails"];
                        apiUrl = $"{baseUrl}{DecryptEWayBillDetails}";

                        response = await httpClient.PostAsync(apiUrl, content);
                        responseData = await response.Content.ReadAsStringAsync();
                        result = JsonConvert.DeserializeObject<JObject>(responseData);
                        //Console.WriteLine($"Response-8 : {result}");
                        headersDict = httpClient.DefaultRequestHeaders.ToDictionary(h => h.Key, h => string.Join(", ", h.Value));
                        headersJson = System.Text.Json.JsonSerializer.Serialize(headersDict);
                        ApiData = new APIsDataModel();
                        {
                            ApiData.ClientGstin = ClientGSTIN;
                            ApiData.RequestNumber = ticketnumber;
                            ApiData.SessionID = sessionId;
                            ApiData.RequestURL = apiUrl;
                            ApiData.RequestParameters = parameters;
                            ApiData.RequestHeaders = headersJson;
                            ApiData.RequestBody = JsonConvert.SerializeObject(RequestBody8);
                            ApiData.Response = responseData;
                            ApiData.ResponseCode = $"{(int)response.StatusCode} {response.StatusCode}";
                            ApiData.Status = getStatus((int)response.StatusCode);
                        }
                        ;
                        await _sLEWayBillBusiness.saveEwayBillAPIsData(ApiData);
                        //Console.WriteLine($"ewb no : {result["ewbNo"]}");
                        //Console.WriteLine($"ewb no : {result["decrypted_data"]["ewbNo"]}");					
                        if (!result.ContainsKey("decrypted_data") || string.IsNullOrEmpty(result["decrypted_data"]?.ToString()))
                        {
                            throw new Exception("8th API Call - Failed to retrieve 'decrypted_data' from the response.");
                        }
                        // Convert decrypted data to DataTable
                        ConvertJsonToDataTable(result, EWBDataTable);
                        //Console.ReadKey();
                        #endregion
                    }
                    #endregion
                }

                //Console.WriteLine("Ewb s" + i);
                await _sLEWayBillBusiness.SaveEWayBillDataAsync(EWBDataTable, ticketnumber, ClientGSTIN);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"{ex.Message} ";
                return View("/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedgerGSTAPI_IRIS.cshtml");
            }

            #endregion

            //Console.WriteLine("Count" + i);
            //Console.WriteLine("Press any key to continue...");
            //Console.ReadKey();

            // Comparison logic (unchanged)
            DataTable SLInvoice = await _sLDataBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            DataTable SLWayBill = await _sLEWayBillBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            DataTable SLEInvoice = await _sLEInvoiceBusiness.GetUserDataBasedOnTicketAsync(ticketnumber, ClientGSTIN);
            //Console.WriteLine("SLInvoice Columns: " + string.Join(", ", SLInvoice.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
            //Console.WriteLine("SLWayBill Columns: " + string.Join(", ", SLWayBill.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
            //Console.WriteLine("SLEInvoice Columns: " + string.Join(", ", SLEInvoice.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));

            string EInvoice_AdminFileName = $"{ticketnumber}_EInvoice_IRIS_API.csv";
            string EWayBill_AdminFileName = $"{ticketnumber}_EWayBill_IRIS_API.csv";

            await _sLComparedDataBusiness.CompareDataAsync(SLInvoice, SLWayBill, SLEInvoice);
            await _sLTicketsBusiness.UpdateSLTicketAsync(ticketnumber, ClientGSTIN, EInvoice_AdminFileName, EWayBill_AdminFileName);

            List<string> catagorys = new List<string>();
            var SLMatchTypes = _configuration["SLMatchTypes"].Split(',').Select(x => x.Trim()).ToList();
            //int i = 0;
            //foreach (var matchType in SLMatchTypes)
            //{
            //   Console.WriteLine($"Match Type {i}: {matchType}");
            //    i++;
            //}
            catagorys.Add(SLMatchTypes[8]);
            catagorys.Add(SLMatchTypes[9]);
            catagorys.Add(SLMatchTypes[10]);
            catagorys.Add(SLMatchTypes[11]);

            // remove that data (combination of category and invoice date) from data

            var data = await _sLComparedDataBusiness.GetComparedData_API(ticketnumber, ClientGSTIN, catagorys, otherdays, formate);
            ViewBag.ReportDataList = data;
            //_logger.LogInformation("Total Compared Data Rows: " + data.Count);
            // store number and sum in a model and save it in 
            ViewBag.Summary = GenerateSLSummary(data);
            ViewBag.GrandTotal = getGrandTotal(data);
			ViewBag.MatchType = SLMatchTypes;
            //var summary = ViewBag.Summary;
            // Console.WriteLine($"Summary : {summary.catagory1InvoiceSum}");
            return View("~/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedger.cshtml");
        }

        public static string getStatus(int statusCode)
        {
            if (statusCode >= 200 && statusCode < 300)
            {
                return "Success Response";
            }
            else if (statusCode >= 400 && statusCode < 500)
            {
                return "Client Error";
            }
            else if (statusCode >= 500 && statusCode < 600)
            {
                return "Server Error";
            }
            else
            {
                return "Other Response (e.g., Redirection or Informational)";
            }
        }

        public static List<string> GetEwbNos(string json)
        {
            var ewbNos = new List<string>();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement.GetProperty("decrypted_data");

            foreach (var item in root.EnumerateArray())
            {
                if (item.TryGetProperty("ewbNo", out var ewbNoProp))
                {
                    ewbNos.Add(ewbNoProp.GetRawText().Trim('"'));
                }
            }

            return ewbNos;
        }

        public void ConvertJsonToDataTable(JObject result, DataTable EWBDataTable)
        {
            string[] allowedFormats = _configuration["Date Format"].Split(',').Select(s => s.Trim()).ToArray();
            var decryptedData = result["decrypted_data"] as JObject;
            DataRow row = EWBDataTable.NewRow();

            row["EWB No"] = decryptedData["ewbNo"]?.ToString();

            string dateString = decryptedData["ewayBillDate"]?.ToString();
            DateTime ewbDate = DateTime.ParseExact(dateString, allowedFormats, CultureInfo.InvariantCulture);
            row["EWB Date"] = ewbDate.ToString("dd-MM-yyyy hh:mm:ss tt");

            row["Supply Type"] = decryptedData["supplyType"]?.ToString();
            row["Doc.No"] = decryptedData["docNo"]?.ToString();

            string dateString2 = decryptedData["docDate"]?.ToString();
            DateTime docDate = DateTime.ParseExact(dateString2, allowedFormats, CultureInfo.InvariantCulture);
            row["Doc.Date"] = docDate.ToString("dd-MM-yyyy hh:mm:ss tt");

            row["Doc.Type"] = decryptedData["docType"]?.ToString();
            row["TO GSTIN"] = decryptedData["toGstin"]?.ToString();
            row["status"] = decryptedData["status"]?.ToString();
            row["No of Items"] = decryptedData["itemList"]?.Count() ?? 0;

            // Get first item for HSN Code and Description
            var firstItem = decryptedData["itemList"]?.First;
            row["Main HSN Code"] = firstItem?["hsnCode"]?.ToString();
            row["Main HSN Desc"] = firstItem?["productDesc"]?.ToString();

            // Sum of taxable amounts
            double assessableValue = 0;
            foreach (var item in decryptedData["itemList"])
            {
                assessableValue += (double?)item["taxableAmount"] ?? 0;

            }
            row["Assessable Value"] = assessableValue;
            row["SGST Value"] = decryptedData["sgstValue"]?.ToString();
            row["CGST Value"] = decryptedData["cgstValue"]?.ToString();
            //row["IGST Value"] = decryptedData["igstValue"]?.ToString();
            //row["CESS Value"] = decryptedData["cessValue"]?.ToString();
            row["Total Invoice Value"] = decryptedData["totInvValue"]?.ToString();

            string dateString3 = decryptedData["validUpto"]?.ToString();
            DateTime validTillDate = DateTime.ParseExact(dateString3, allowedFormats, CultureInfo.InvariantCulture);
            row["Valid Till Date"] = validTillDate.ToString("dd-MM-yyyy hh:mm:ss tt");

            row["Gen.Mode"] = decryptedData["genMode"]?.ToString();



            //var vehicle = decryptedData["VehiclListDetails"]?.First;
            //row["Transporter Details"] = vehicle?["vehicleNo"]?.ToString() ?? "";

            //row["From GSTIN"] = decryptedData["fromGstin"]?.ToString();			
            //row["From GSTIN Info"] = $"{decryptedData["fromTrdName"]}, {decryptedData["fromAddr1"]}";
            //row["TO GSTIN Info"] = $"{decryptedData["toTrdName"]}, {decryptedData["toAddr1"]}";			
            //row["CESS Non.Advol Value"] = decryptedData["cessNonAdvolValue"]?.ToString();
            //row["Other Value"] = decryptedData["otherValue"]?.ToString();		
            //row["Other Party Rejection Status"] = decryptedData["rejectStatus"]?.ToString();
            //row["IRN"] = ""; // Not present in the JSON


            EWBDataTable.Rows.Add(row);

        }

        public DataTable DecryptEinvoiceResponse(string response, string ek, string userGstin, DataTable EInvoiceDataTable)
        {
            // Decode the key
            byte[] key = Convert.FromBase64String(ek);

            // Test base64 decoding
            Convert.FromBase64String(response);

            // Decrypt the response
            byte[] resul = AesDecryptWithKey(response, key);
            //Console.WriteLine("1");

            // Decode the result to a UTF-8 string
            string jsonStr = Encoding.UTF8.GetString(resul);
            //Console.WriteLine("2");


            // Additional base64 decoding of the decrypted string
            jsonStr = Encoding.UTF8.GetString(Convert.FromBase64String(jsonStr));

            // Try parsing JSON
            JsonDocument jsonData = JsonDocument.Parse(jsonStr);

            string jsonPreview = System.Text.Json.JsonSerializer.Serialize(jsonData, new JsonSerializerOptions { WriteIndented = true });


            var result = JsonConvert.DeserializeObject<JObject>(jsonPreview);

            JArray dataArray = (JArray)result["data"];

            if (dataArray != null)
            {
                foreach (JObject item in dataArray)
                {
                    var signedInvoice = item["SignedInvoice"]?.ToString();
                    if (!string.IsNullOrEmpty(signedInvoice))
                    {
                        ConvertToEInvoiceDataTable(EInvoiceDataTable, signedInvoice);
                    }
                }
            }
            return EInvoiceDataTable;
        }

        public void ConvertToEInvoiceDataTable(DataTable EInvoiceDataTable, string signedInvoice)
        {
            // Create a JwtSecurityTokenHandler to parse the JWT
            var jwtHandler = new JwtSecurityTokenHandler();
            var jwtToken = jwtHandler.ReadJwtToken(signedInvoice);

            // Extract the payload (e-invoice data) from the JWT
            string payloadJson = jwtToken.Payload.SerializeToJson();

            // Parse the payload to extract the "data" field, which contains the e-invoice JSON
            var payloadObject = Newtonsoft.Json.Linq.JObject.Parse(payloadJson);
            string einvoiceData = payloadObject["data"].ToString();

            // Parse the e-invoice data into a JObject for formatting
            var einvoiceObject = Newtonsoft.Json.Linq.JObject.Parse(einvoiceData);
            var output = einvoiceObject.ToString(Newtonsoft.Json.Formatting.Indented);
            // Print the decoded e-invoice data in plain text (formatted JSON)

            var _SLEInvoiceColumns = _configuration["SLEInvoice"].Split(',').Select(x => x.Trim()).ToList();

            var row = EInvoiceDataTable.NewRow();

            var buyer = einvoiceObject["BuyerDtls"];
            var doc = einvoiceObject["DocDtls"];
            var tran = einvoiceObject["TranDtls"];
            var val = einvoiceObject["ValDtls"];
            var itemList = einvoiceObject["ItemList"] as JArray;

            decimal totalValue = 0;
            decimal igstAmt = 0;
            decimal cgstAmt = 0;
            decimal sgstAmt = 0;
            decimal cessAmt = 0;
            string gstRate = "";

            foreach (var item in itemList)
            {
                gstRate = item?["GstRt"]?.ToString() ?? "";

                totalValue += Convert.ToDecimal(item?["AssAmt"]?.ToString() ?? "0");
                igstAmt += Convert.ToDecimal(item?["IgstAmt"]?.ToString() ?? "0");
                cgstAmt += Convert.ToDecimal(item?["CgstAmt"]?.ToString() ?? "0");
                sgstAmt += Convert.ToDecimal(item?["SgstAmt"]?.ToString() ?? "0");
                cessAmt += Convert.ToDecimal(item?["StateCesAmt"]?.ToString() ?? "0");
            }
            //GSTIN/UIN of Recipient,Receiver Name,Invoice Number,Invoice date,Invoice Value,Place Of Supply,
            //Reverse Charge,Applicable % of Tax Rate,Invoice Type,E-Commerce GSTIN,Rate,
            //Taxable Value,Integrated Tax,Central Tax,State/UT Tax,Cess Amount,
            //IRN,IRN date,E-invoice status,GSTR-1 auto-population/ deletion upon cancellation date,GSTR-1 auto-population/ deletion status,
            //Error in auto-population/ deletion"

            row[_SLEInvoiceColumns[0]] = buyer?["Gstin"]?.ToString();
            row[_SLEInvoiceColumns[1]] = buyer?["LglNm"]?.ToString();
            row[_SLEInvoiceColumns[2]] = doc?["No"]?.ToString();
            row[_SLEInvoiceColumns[3]] = doc?["Dt"]?.ToString();
            row[_SLEInvoiceColumns[4]] = val?["TotInvVal"]?.ToString();
            row[_SLEInvoiceColumns[5]] = buyer?["Pos"]?.ToString();
            row[_SLEInvoiceColumns[6]] = tran?["RegRev"]?.ToString();
            row[_SLEInvoiceColumns[7]] = ""; // Not available
            row[_SLEInvoiceColumns[8]] = doc?["Typ"]?.ToString();
            row[_SLEInvoiceColumns[9]] = ""; // Not available
            row[_SLEInvoiceColumns[10]] = gstRate;
            row[_SLEInvoiceColumns[11]] = totalValue.ToString();
            row[_SLEInvoiceColumns[12]] = igstAmt.ToString();
            row[_SLEInvoiceColumns[13]] = cgstAmt.ToString();
            row[_SLEInvoiceColumns[14]] = sgstAmt.ToString();
            row[_SLEInvoiceColumns[15]] = cessAmt.ToString();
            row[_SLEInvoiceColumns[16]] = einvoiceObject?["Irn"]?.ToString();
            row[_SLEInvoiceColumns[17]] = einvoiceObject?["AckDt"]?.ToString();
            row[_SLEInvoiceColumns[18]] = ""; // string.IsNullOrEmpty(einvoiceObject?["Irn"]?.ToString()) ? "Not Generated" : "Generated";
            row[_SLEInvoiceColumns[19]] = ""; // Not available
            row[_SLEInvoiceColumns[20]] = ""; // Not available
            row[_SLEInvoiceColumns[21]] = ""; // Not available

            EInvoiceDataTable.Rows.Add(row);
        }

        #endregion

        #region SalesLedgerClosedRequests
        public async Task<IActionResult> SalesLedgerClosedRequests(DateTime? fromdate, DateTime? todate)
		{
			DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
			DateTime toDateTime = todate ?? DateTime.Now; // Default to today

			//await TjCaptions("CompletedTasks");
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act6";
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                return View("~/Views/Admin/SalesLedgerClosedRequests/SalesLedgerClosedRequests.cshtml");
            }

   //         var email = MySession.Current.Email; // Get the email from session
			//var clients = await _userBusiness.GetAdminClients(email);
			//string[] gstinArray = clients.Select(c => c.ClientGSTIN).ToArray();
			//var Tickets = await _sLTicketsBusiness.GetClientsClosedTicketsStatusAsync(gstinArray, fromDateTime, toDateTime);

            var Tickets = await _sLTicketsBusiness.GetCloseTicketsStatusAsync(MySession.Current.gstin, fromDateTime, toDateTime); //(ViewBag.UserGSTIN); // Fetch pending tickets based on user_gstin

            ViewBag.Tickets = Tickets;

			return View("~/Views/Admin/SalesLedgerClosedRequests/SalesLedgerClosedRequests.cshtml");
		}

		public async Task<IActionResult> CloseSLRequest(string RequestNumber, string ClientGstin)
		{
			string requestNo = RequestNumber;
			string ClientGSTIN = ClientGstin;
			//await TjCaptions("CloseRequest");
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act6"; // For active button styling
			await _sLTicketsBusiness.CloseSLTicketAsync(requestNo, ClientGSTIN);

			var ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			string toEmail = _configuration["Mail:ToMail"];
			if (_configuration["Mail:SendToClient"] == "Yes")
			{
				toEmail = ticket.CLientEmail;
			}
			string fromMail = _configuration["Mail:FromMail"];

			string subjectTemplate = _configuration["Mail:SubjectTemplate"];
			string bodyTemplate = _configuration["Mail:BodyTemplate"];

			string type = _configuration["Mail:Type2"];

			string subject = subjectTemplate
				.Replace("{Type}", type)
				.Replace("{RequestNo}", requestNo);

			Attachment attachment1 = await GenerateSLInvoiceExcelAttachmentAsync(requestNo, ClientGSTIN);
			Attachment attachment2 = await GenerateSLExcelAttachmentAsync(requestNo, ticket.FileName, ClientGSTIN);


			string body = bodyTemplate
				.Replace("{CustomerName}", ticket.ClientName)
				.Replace("{RequestNo}", requestNo)
				.Replace("{CreatedDate}", ticket.RequestCreatedDate?.ToString("yyyy-MMM-dd hh:mm:ss tt"))
				.Replace("{ClosedDate}", DateTime.Now.ToString("yyyy-MMM-dd hh:mm:ss tt"))
				.Replace("{FileName}", ticket.FileName)
				.Replace("{OutputFileName}", attachment2.Name);


			// Send email
			string[] ccList = _configuration["Mail:CCMail"].Split(',');

			await SendEmailAsync(
				toEmail,
				subject,
				body,
				ccList,
				attachment1,
				attachment2
			);


			return Json(new { success = true });
		}

		#endregion

		#region Sales Register Export XL
		public async Task<IActionResult> ExportSLInvoiceFile(string requestNo, string ClientGSTIN)
		{

			var ticketDetails = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			var data = await _sLDataBusiness.GetSLInvoiceData(requestNo, ClientGSTIN);
			data = data.OrderBy(x => Convert.ToInt32(x.Sno)).ToList();
		   
			var fileName = ticketDetails.FileName;
			string[] parts = fileName.Split('.');
			string fileExtension = parts[1].ToLower();

			var _SLinvoice = _configuration["SLInvoice"];
			string[] headers = _SLinvoice.Split(',').Select(x => x.Trim()).ToArray();

			if (fileExtension == "csv")
			{
				var csvBuilder = new StringBuilder();
				// Write headers
				csvBuilder.AppendLine(string.Join(",", headers));

				// Write data
				foreach (var item in data)
				{ //S.NO	INV NO	DATE	NAME OF CUSTOMERS	GST No	State to Supply(POS)	
                  //TAXABLE AMOUNT	IGST	SGST	CGST	TOTAL	GST Rates	HSN/SAC Code	
                  //Qty	Units	Material Description	Type of Invoice  User Gstin IRN

                    var row = new string[]
					{
						item.Sno,
						item.InvoiceNumber,
						item.InvoiceDate?.ToString("yyyy-MM-dd") ?? "", // format date
						item.CustomerName,
						item.CustomerGSTIN,
						item.StateToSupply,
						item.TaxableAmount.ToString(),
						item.IGST.ToString(),
						item.SGST.ToString(),
						item.CGST.ToString(),
						item.TotalAmount.ToString(),
						item.GSTRate.ToString(),
						item.HSNSACCode,
						item.Quantity.ToString(),
						item.Units,
						item.MaterialDescription,
						item.TypeOfInvoice,
                        ClientGSTIN,
						item.Irn

					};

					csvBuilder.AppendLine(string.Join(",", row.Select(field => EscapeCsv(field))));
				}

				byte[] content = Encoding.UTF8.GetBytes(csvBuilder.ToString());

				return File(content, "text/csv", fileName);
			}
			if (fileExtension == "xlsx")
			{
				using var workbook = new XLWorkbook();
				var worksheet = workbook.Worksheets.Add("Sales Register");

				for (int i = 0; i < headers.Length; i++)
				{
					worksheet.Cell(1, i + 1).Value = headers[i];
					worksheet.Cell(1, i + 1).Style.Font.Bold = true;
				}

				int rowIndex = 2;
				foreach (var item in data)
				{
					worksheet.Cell(rowIndex, 1).Value = item.Sno;
					worksheet.Cell(rowIndex, 2).Value = item.InvoiceNumber;
					worksheet.Cell(rowIndex, 3).Value = item.InvoiceDate?.ToString("yyyy-MM-dd");
					worksheet.Cell(rowIndex, 4).Value = item.CustomerName;
					worksheet.Cell(rowIndex, 5).Value = item.CustomerGSTIN;
					worksheet.Cell(rowIndex, 6).Value = item.StateToSupply;
					worksheet.Cell(rowIndex, 7).Value = item.TaxableAmount;
					worksheet.Cell(rowIndex, 8).Value = item.IGST;
					worksheet.Cell(rowIndex, 9).Value = item.SGST;
					worksheet.Cell(rowIndex, 10).Value = item.CGST;
					worksheet.Cell(rowIndex, 11).Value = item.TotalAmount;
					worksheet.Cell(rowIndex, 12).Value = item.GSTRate;
					worksheet.Cell(rowIndex, 13).Value = item.HSNSACCode;
					worksheet.Cell(rowIndex, 14).Value = item.Quantity;
					worksheet.Cell(rowIndex, 15).Value = item.Units;
					worksheet.Cell(rowIndex, 16).Value = item.MaterialDescription;
					worksheet.Cell(rowIndex, 17).Value = item.TypeOfInvoice;
					worksheet.Cell(rowIndex, 18).Value = ClientGSTIN;
					worksheet.Cell(rowIndex, 19).Value = item.Irn;
					rowIndex++;
				}

				var stream = new MemoryStream();
				workbook.SaveAs(stream);
				stream.Position = 0; // Reset stream position
				return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",fileName);


			}
			if (fileExtension == "xls")
			{
				HSSFWorkbook workbook = new HSSFWorkbook();
				ISheet sheet = workbook.CreateSheet("Sales Register");

				var headerRow = sheet.CreateRow(0);
				for (int i = 0; i < headers.Length; i++)
				{
					headerRow.CreateCell(i).SetCellValue(headers[i]);
				}

				int rowIndex = 1;
				foreach (var item in data)
				{
					var row = sheet.CreateRow(rowIndex++);
					row.CreateCell(0).SetCellValue(item.Sno);
					row.CreateCell(1).SetCellValue(item.InvoiceNumber);
					row.CreateCell(2).SetCellValue(item.InvoiceDate?.ToString("yyyy-MM-dd") ?? "");
					row.CreateCell(3).SetCellValue(item.CustomerName);
					row.CreateCell(4).SetCellValue(item.CustomerGSTIN);
					row.CreateCell(5).SetCellValue(item.StateToSupply);
					row.CreateCell(6).SetCellValue((double)item.TaxableAmount);
					row.CreateCell(7).SetCellValue((double)item.IGST);
					row.CreateCell(8).SetCellValue((double)item.SGST);
					row.CreateCell(9).SetCellValue((double)item.CGST);
					row.CreateCell(10).SetCellValue((double)item.TotalAmount);
					row.CreateCell(11).SetCellValue(item.GSTRate);
					row.CreateCell(12).SetCellValue(item.HSNSACCode);
					row.CreateCell(13).SetCellValue(item.Quantity);
					row.CreateCell(14).SetCellValue(item.Units);
					row.CreateCell(15).SetCellValue(item.MaterialDescription);
					row.CreateCell(16).SetCellValue(item.TypeOfInvoice);
					row.CreateCell(17).SetCellValue(ClientGSTIN);
					row.CreateCell(18).SetCellValue(item.Irn);
				}
				var stream = new MemoryStream();
				workbook.Write(stream);
				stream.Position = 0; // Reset stream position
				return File(stream.ToArray(), "application/vnd.ms-excel", fileName);
			}
			return BadRequest("Unsupported file format.");
		}
		public async Task<Attachment> GenerateSLInvoiceExcelAttachmentAsync(string requestNo, string ClientGSTIN)
		{
			var ticketDetails = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			var data = await _sLDataBusiness.GetSLInvoiceData(requestNo, ClientGSTIN);
			data = data.OrderBy(x => Convert.ToInt32(x.Sno)).ToList();
			var fileName = ticketDetails.FileName;
			string[] parts = fileName.Split('.');
			string fileExtension = parts[1].ToLower();

			var _SLinvoice = _configuration["SLInvoice"];
			string[] headers = _SLinvoice.Split(',').Select(x => x.Trim()).ToArray();

			Attachment attachment = null!;

			if (fileExtension == "csv")
			{
				var csvBuilder = new StringBuilder();
				// Write headers
				csvBuilder.AppendLine(string.Join(",", headers));
				// Write data
				foreach (var item in data)
				{ //S.NO	INV NO	DATE	NAME OF CUSTOMERS	GST No	State to Supply(POS)	
				  //TAXABLE AMOUNT	IGST	SGST	CGST	TOTAL	GST Rates	HSN/SAC Code	
				  //Qty	Units	Material Description	Type of Invoice
					var row = new string[]
					{
						 item.Sno,
						 item.InvoiceNumber,
						 item.InvoiceDate?.ToString("yyyy-MM-dd") ?? "", // format date
						 item.CustomerName,
						 item.CustomerGSTIN,
						 item.StateToSupply,
						 item.TaxableAmount.ToString(),
						 item.IGST.ToString(),
						 item.SGST.ToString(),
						 item.CGST.ToString(),
						 item.TotalAmount.ToString(),
						 item.GSTRate.ToString(),
						 item.HSNSACCode,
						 item.Quantity.ToString(),
						 item.Units,
						 item.MaterialDescription,
						 item.TypeOfInvoice,
                         ClientGSTIN,
						 item.Irn

                    };

					csvBuilder.AppendLine(string.Join(",", row.Select(field => EscapeCsv(field))));
				}

				byte[] content = Encoding.UTF8.GetBytes(csvBuilder.ToString());
				var stream = new MemoryStream(content);
				stream.Position = 0; // Reset stream before using
				attachment = new Attachment(stream, fileName, "text/csv");
			}
			if (fileExtension == "xlsx")
			{
				using var workbook = new XLWorkbook();
				var worksheet = workbook.Worksheets.Add("SL Invoices");

				for (int i = 0; i < headers.Length; i++)
				{
					worksheet.Cell(1, i + 1).Value = headers[i];
					worksheet.Cell(1, i + 1).Style.Font.Bold = true;
				}

				int rowIndex = 2;
				foreach (var item in data)
				{
					worksheet.Cell(rowIndex, 1).Value = item.Sno;
					worksheet.Cell(rowIndex, 2).Value = item.InvoiceNumber;
					worksheet.Cell(rowIndex, 3).Value = item.InvoiceDate?.ToString("yyyy-MM-dd");
					worksheet.Cell(rowIndex, 4).Value = item.CustomerName;
					worksheet.Cell(rowIndex, 5).Value = item.CustomerGSTIN;
					worksheet.Cell(rowIndex, 6).Value = item.StateToSupply;
					worksheet.Cell(rowIndex, 7).Value = item.TaxableAmount;
					worksheet.Cell(rowIndex, 8).Value = item.IGST;
					worksheet.Cell(rowIndex, 9).Value = item.SGST;
					worksheet.Cell(rowIndex, 10).Value = item.CGST;
					worksheet.Cell(rowIndex, 11).Value = item.TotalAmount;
					worksheet.Cell(rowIndex, 12).Value = item.GSTRate;
					worksheet.Cell(rowIndex, 13).Value = item.HSNSACCode;
					worksheet.Cell(rowIndex, 14).Value = item.Quantity;
					worksheet.Cell(rowIndex, 15).Value = item.Units;
					worksheet.Cell(rowIndex, 16).Value = item.MaterialDescription;
					worksheet.Cell(rowIndex, 17).Value = item.TypeOfInvoice;
					worksheet.Cell(rowIndex, 18).Value = ClientGSTIN;
					worksheet.Cell(rowIndex, 19).Value = item.Irn;

					rowIndex++;
				}
				var stream = new MemoryStream();
				workbook.SaveAs(stream);
				stream.Position = 0;
				attachment = new Attachment(stream, $"{fileName}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
			}
			if (fileExtension == "xls")
			{
				HSSFWorkbook workbook = new HSSFWorkbook();
				ISheet sheet = workbook.CreateSheet("SL Invoices");

				var headerRow = sheet.CreateRow(0);
				for (int i = 0; i < headers.Length; i++)
				{
					headerRow.CreateCell(i).SetCellValue(headers[i]);
				}

				int rowIndex = 1;
				foreach (var item in data)
				{
					var row = sheet.CreateRow(rowIndex++);
					row.CreateCell(0).SetCellValue(item.Sno);
					row.CreateCell(1).SetCellValue(item.InvoiceNumber);
					row.CreateCell(2).SetCellValue(item.InvoiceDate?.ToString("yyyy-MM-dd") ?? "");
					row.CreateCell(3).SetCellValue(item.CustomerName);
					row.CreateCell(4).SetCellValue(item.CustomerGSTIN);
					row.CreateCell(5).SetCellValue(item.StateToSupply);
					row.CreateCell(6).SetCellValue((double)item.TaxableAmount);
					row.CreateCell(7).SetCellValue((double)item.IGST);
					row.CreateCell(8).SetCellValue((double)item.SGST);
					row.CreateCell(9).SetCellValue((double)item.CGST);
					row.CreateCell(10).SetCellValue((double)item.TotalAmount);
					row.CreateCell(11).SetCellValue(item.GSTRate);
					row.CreateCell(12).SetCellValue(item.HSNSACCode);
					row.CreateCell(13).SetCellValue(item.Quantity);
					row.CreateCell(14).SetCellValue(item.Units);
					row.CreateCell(15).SetCellValue(item.MaterialDescription);
					row.CreateCell(16).SetCellValue(item.TypeOfInvoice);
					row.CreateCell(17).SetCellValue(ClientGSTIN);
					row.CreateCell(18).SetCellValue(item.Irn);
				}

				var stream = new MemoryStream();
				workbook.Write(stream);
				stream.Position = 0;
				//return File(stream.ToArray(), "application/vnd.ms-excel", $"{fileName}.xls");
				attachment = new Attachment(stream, $"{fileName}.xls", "application/vnd.ms-excel");

			}

			return attachment;
		}
		public async Task<IActionResult> ExportSLEWayBillFile(string requestNo, string ClientGSTIN)
		{
			var ticketDetails = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			var data = await _sLEWayBillBusiness.GetSLEWayBillData(requestNo, ClientGSTIN);
			//data = data.OrderBy(x => Convert.ToInt32(x.Sno)).ToList();
			
			var fileName = ticketDetails.EWayBillFileName;
			string[] parts = fileName.Split('.');
			string fileExtension = parts[1].ToLower();
		   
			var EWayBill = _configuration["SLEWayBill"];
			string[] headers = EWayBill.Split(',').Select(x => x.Trim()).ToArray();

			if (fileExtension == "csv")
			{
				var csvBuilder = new StringBuilder();
				//Write headers
				csvBuilder.AppendLine(string.Join(",", headers));
				//EWB No	EWB Date	Supply Type	Doc.No	Doc.Date	Doc.Type	TO GSTIN	status	No of Items	
				//Main HSN Code	Main HSN Desc	Assessable Value	SGST Value	CGST Value	Total Invoice Value	Valid Till Date	Gen.Mode
				foreach (var item in data)
				{
					var row = new string[]
					{
						item.EWBNumber,
						item.EWBDate?.ToString("yyyy-MM-dd") ?? "", // format date
						item.SupplyType,
						item.DocNumber,
						item.DocDate?.ToString("yyyy-MM-dd") ?? "", // format date
						item.DocType,
						item.TOGSTIN,
						item.status,
						item.NoofItems.ToString(),
						item.MainHSNCode,
						item.MainHSNDesc,
						item.AssessableValue.ToString(),
						item.SGST.ToString(),
						item.CGST.ToString(),
						item.IGST.ToString(),
						item.CESS.ToString(),
						item.TotalInvoiceValue.ToString(),
						item.ValidTillDate?.ToString("yyyy-MM-dd") ?? "", // format date
						item.GenMode
					};
					csvBuilder.AppendLine(string.Join(",", row.Select(field => EscapeCsv(field))));
				}
				byte[] content = Encoding.UTF8.GetBytes(csvBuilder.ToString());
				return File(content, "text/csv", fileName);
			}
			if (fileExtension == "xlsx")
			{
				using var workbook = new XLWorkbook();
				var worksheet = workbook.Worksheets.Add("EWayBill");

				for (int i = 0; i < headers.Length; i++)
				{
					worksheet.Cell(1, i + 1).Value = headers[i];
					worksheet.Cell(1, i + 1).Style.Font.Bold = true;
				}

				int rowIndex = 2;
				foreach (var item in data)
				{
					worksheet.Cell(rowIndex, 1).Value = item.EWBNumber; 
					worksheet.Cell(rowIndex, 2).Value = item.EWBDate?.ToString("yyyy-MM-dd");
					worksheet.Cell(rowIndex, 3).Value = item.SupplyType;
					worksheet.Cell(rowIndex, 4).Value = item.DocNumber;
					worksheet.Cell(rowIndex, 5).Value = item.DocDate?.ToString("yyyy-MM-dd");
					worksheet.Cell(rowIndex, 6).Value = item.DocType;
					worksheet.Cell(rowIndex, 7).Value = item.TOGSTIN;
					worksheet.Cell(rowIndex, 8).Value = item.status;
					worksheet.Cell(rowIndex, 9).Value = item.NoofItems;
					worksheet.Cell(rowIndex, 10).Value = item.MainHSNCode;
					worksheet.Cell(rowIndex, 11).Value = item.MainHSNDesc;
					worksheet.Cell(rowIndex, 12).Value = item.AssessableValue;
					worksheet.Cell(rowIndex, 13).Value = item.SGST;
					worksheet.Cell(rowIndex, 14).Value = item.CGST;
					worksheet.Cell(rowIndex, 15).Value = item.IGST;
					worksheet.Cell(rowIndex, 16).Value = item.CESS;
					worksheet.Cell(rowIndex, 17).Value = item.TotalInvoiceValue;
					worksheet.Cell(rowIndex, 18).Value = item.ValidTillDate?.ToString("yyyy-MM-dd");
					worksheet.Cell(rowIndex, 19).Value = item.GenMode;                    
					rowIndex++;
				}
				var stream = new MemoryStream();
				workbook.SaveAs(stream);
				stream.Position = 0; // Reset stream position
				return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",fileName);
			}
			if (fileExtension == "xls")
			{
				HSSFWorkbook workbook = new HSSFWorkbook();
				ISheet sheet = workbook.CreateSheet("EWayBill");
				var headerRow = sheet.CreateRow(0);
				for (int i = 0; i < headers.Length; i++)
				{
					headerRow.CreateCell(i).SetCellValue(headers[i]);
				}
				int rowIndex = 1;
				foreach (var item in data)
				{
					var row = sheet.CreateRow(rowIndex++);
					row.CreateCell(0).SetCellValue(item.EWBNumber);
					row.CreateCell(1).SetCellValue(item.EWBDate?.ToString("yyyy-MM-dd") ?? "");
					row.CreateCell(2).SetCellValue(item.SupplyType);
					row.CreateCell(3).SetCellValue(item.DocNumber);
					row.CreateCell(4).SetCellValue(item.DocDate?.ToString("yyyy-MM-dd") ?? "");
					row.CreateCell(5).SetCellValue(item.DocType);
					row.CreateCell(6).SetCellValue(item.TOGSTIN);
					row.CreateCell(7).SetCellValue(item.status);
					row.CreateCell(8).SetCellValue(item.NoofItems);
					row.CreateCell(9).SetCellValue(item.MainHSNCode);
					row.CreateCell(10).SetCellValue(item.MainHSNDesc);
					row.CreateCell(11).SetCellValue((double)item.AssessableValue);
					row.CreateCell(12).SetCellValue((double)item.SGST);
					row.CreateCell(13).SetCellValue((double)item.CGST);
					row.CreateCell(14).SetCellValue((double)item.IGST);
					row.CreateCell(15).SetCellValue((double)item.CESS);
					row.CreateCell(16).SetCellValue((double)item.TotalInvoiceValue);
					row.CreateCell(17).SetCellValue(item.ValidTillDate?.ToString("yyyy-MM-dd") ?? "");
					row.CreateCell(18).SetCellValue(item.GenMode);                
				}
				var stream = new MemoryStream();
				workbook.Write(stream);
				stream.Position = 0; // Reset stream position
				return File(stream.ToArray(), "application/vnd.ms-excel", fileName);
			}

			return BadRequest("Unsupported file format.");
		}
		public async Task<IActionResult> ExportSLEInvoiceFile(string requestNo, string ClientGSTIN)
		{
			var ticketDetails = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			var data = await _sLEInvoiceBusiness.GetSLEInvoiceData(requestNo, ClientGSTIN);
			//data = data.OrderBy(x => Convert.ToInt32(x.Sno)).ToList();
			
			var fileName = ticketDetails.EInvoiceFileName;
			string[] parts = fileName.Split('.');
			string fileExtension = parts[1];

			var EInvoice = _configuration["SLEInvoice"];
			string[] headers = EInvoice.Split(',').Select(x => x.Trim()).ToArray();

			if(fileExtension == "csv")
			{
				var csvBuilder = new StringBuilder();
				//Write headers
				csvBuilder.AppendLine(string.Join(",", headers));
				//GSTIN/UIN of Recipient,Receiver Name,Invoice Number,Invoice date,Invoice Value,Place Of Supply,
				//Reverse Charge,Applicable % of Tax Rate,Invoice Type,E-Commerce GSTIN,Rate,
				//Taxable Value,Integrated Tax,Central Tax,State/UT Tax,Cess Amount,
				//IRN,IRN date,E-invoice status,GSTR-1 auto-population/ deletion upon cancellation date,
				//GSTR-1 auto-population/ deletion status,Error in auto-population/ deletion"
				foreach (var item in data)
				{
					var row = new string[]
					{
						  item.RecipientGSTIN,
						  item.ReceiverName,
						  item.InvoiceNumber,
						  item.Invoicedate?.ToString("yyyy-MM-dd") ?? "", // format date
						  item.InvoiceValue.ToString(),
						  item.PlaceOfSupply,
						  item.ReverseCharge.ToString(),
						  item.ApplicableTaxRate.ToString(),
						  item.InvoiceType,
						  item.ECommerceGSTIN,
						  item.TaxRate.ToString(),
						  item.TaxableValue.ToString(),
						  item.IGST.ToString(),
						  item.CGST.ToString(),
						  item.SGST.ToString(),
						  item.CESS.ToString(),
						  item.IRN,
						  item.IRNDate?.ToString("yyyy-MM-dd") ?? "", // format date
						  item.Einvoicestatus,
						  item.GSTR1AutoPopulationDate?.ToString("yyyy-MM-dd") ?? "", // format date
						  item.GSTR1AutoPopulationStatus,
						  item.ErrorInAutoPopulationDeletion
					};
					csvBuilder.AppendLine(string.Join(",", row.Select(field => EscapeCsv(field))));
				}
				byte[] content = Encoding.UTF8.GetBytes(csvBuilder.ToString());
				return File(content, "text/csv", fileName);
			}
			if(fileExtension == "xlsx")
			{
				using var workbook = new XLWorkbook();
				var worksheet = workbook.Worksheets.Add("EInvoice");
				for (int i = 0; i < headers.Length; i++)
				{
					worksheet.Cell(1, i + 1).Value = headers[i];
					worksheet.Cell(1, i + 1).Style.Font.Bold = true;
				}
				int rowIndex = 2;
				foreach (var item in data)
				{
					worksheet.Cell(rowIndex, 1).Value = item.RecipientGSTIN;
					worksheet.Cell(rowIndex, 2).Value = item.ReceiverName;
					worksheet.Cell(rowIndex, 3).Value = item.InvoiceNumber;
					worksheet.Cell(rowIndex, 4).Value = item.Invoicedate?.ToString("yyyy-MM-dd");
					worksheet.Cell(rowIndex, 5).Value = item.InvoiceValue;
					worksheet.Cell(rowIndex, 6).Value = item.PlaceOfSupply;
					worksheet.Cell(rowIndex, 7).Value = item.ReverseCharge;
					worksheet.Cell(rowIndex, 8).Value = item.ApplicableTaxRate;
					worksheet.Cell(rowIndex, 9).Value = item.InvoiceType;
					worksheet.Cell(rowIndex, 10).Value = item.ECommerceGSTIN;
					worksheet.Cell(rowIndex, 11).Value = item.TaxRate;
					worksheet.Cell(rowIndex, 12).Value = item.TaxableValue;
					worksheet.Cell(rowIndex, 13).Value = item.IGST;
					worksheet.Cell(rowIndex, 14).Value = item.CGST;
					worksheet.Cell(rowIndex, 15).Value = item.SGST;
					worksheet.Cell(rowIndex, 16).Value = item.CESS;
					worksheet.Cell(rowIndex, 17).Value = item.IRN;
					worksheet.Cell(rowIndex, 18).Value = item.IRNDate?.ToString("yyyy-MM-dd");
					worksheet.Cell(rowIndex, 19).Value = item.Einvoicestatus;
					worksheet.Cell(rowIndex, 20).Value = item.GSTR1AutoPopulationDate?.ToString("yyyy-MM-dd");
					worksheet.Cell(rowIndex, 21).Value = item.GSTR1AutoPopulationStatus;
					worksheet.Cell(rowIndex, 22).Value = item.ErrorInAutoPopulationDeletion;

					rowIndex++;
				}
				var stream = new MemoryStream();
				workbook.SaveAs(stream);
				stream.Position = 0; // Reset stream position
				return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
			}
			if (fileExtension == "xls")
			{
				HSSFWorkbook workbook = new HSSFWorkbook();
				ISheet sheet = workbook.CreateSheet("EInvoice");
				var headerRow = sheet.CreateRow(0);
				for (int i = 0; i < headers.Length; i++)
				{
					headerRow.CreateCell(i).SetCellValue(headers[i]);
				}
				int rowIndex = 1;
				foreach (var item in data)
				{
					var row = sheet.CreateRow(rowIndex++);
					row.CreateCell(0).SetCellValue(item.RecipientGSTIN);
					row.CreateCell(1).SetCellValue(item.ReceiverName);
					row.CreateCell(2).SetCellValue(item.InvoiceNumber);
					row.CreateCell(3).SetCellValue(item.Invoicedate?.ToString("yyyy-MM-dd") ?? "");
					row.CreateCell(4).SetCellValue((double)item.InvoiceValue);
					row.CreateCell(5).SetCellValue(item.PlaceOfSupply);
					row.CreateCell(6).SetCellValue(item.ReverseCharge);
					row.CreateCell(7).SetCellValue(item.ApplicableTaxRate);
					row.CreateCell(8).SetCellValue(item.InvoiceType);
					row.CreateCell(9).SetCellValue(item.ECommerceGSTIN);
					row.CreateCell(10).SetCellValue(item.TaxRate);
					row.CreateCell(11).SetCellValue((double)item.TaxableValue);
					row.CreateCell(12).SetCellValue((double)item.IGST);
					row.CreateCell(13).SetCellValue((double)item.CGST);
					row.CreateCell(14).SetCellValue((double)item.SGST);
					row.CreateCell(15).SetCellValue((double)item.CESS);
					row.CreateCell(16).SetCellValue(item.IRN);
					row.CreateCell(17).SetCellValue(item.IRNDate?.ToString("yyyy-MM-dd") ?? "");
					row.CreateCell(18).SetCellValue(item.Einvoicestatus);
					row.CreateCell(19).SetCellValue(item.GSTR1AutoPopulationDate?.ToString("yyyy-MM-dd") ?? "");
					row.CreateCell(20).SetCellValue(item.GSTR1AutoPopulationStatus);
					row.CreateCell(21).SetCellValue(item.ErrorInAutoPopulationDeletion);
				}
				var stream = new MemoryStream();
				workbook.Write(stream);
				stream.Position = 0; // Reset stream position
				return File(stream.ToArray(), "application/vnd.ms-excel", fileName);


			}
				
			return BadRequest("Unsupported file format.");
		}

		public async Task<IActionResult> ExportSLReport(string requestNo, string ClientGSTIN)
		{
			var ticketDetails = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			var fileName = ticketDetails.FileName;
			var EWayBillFileName = ticketDetails.EWayBillFileName;
			var EInvoiveFileName = ticketDetails.EInvoiceFileName;

			var status = ticketDetails.TicketStatus;
			string ExportFileName = $"{requestNo}_Report.xlsx";
			//ViewBag.Message = "yes";
			if (status == "Completed")
			{
				ExportFileName = $"{fileName.Split('.')[0]}_{requestNo}_Report.xlsx";
			}
			if (status == "Analysed")
			{
				ExportFileName = $"{fileName.Split('.')[0]}_VS_{EWayBillFileName.Split('.')[0]}_VS_{EInvoiveFileName.Split('.')[0]}_Report.xlsx";
			}


			var data = await _sLComparedDataBusiness.GetComparedDataBasedOnTicketAsync(requestNo, ClientGSTIN);

			var invoiceData = data
				.Where(d => d.DataSource != null && d.DataSource.Trim().Equals("Invoice", StringComparison.OrdinalIgnoreCase))
				.OrderBy(x => x.SNo)
				.ToList();
			//_logger.LogInformation($"Invoice Data Count: {invoiceData.Count}");

			var eWayBillData = data
				.Where(d => d.DataSource != null && d.DataSource.Trim().Equals("EWayBill", StringComparison.OrdinalIgnoreCase))
				.OrderBy(x => x.SNo)
				.ToList();

			var eInvoiceData = data
			   .Where(d => d.DataSource != null && d.DataSource.Trim().Equals("EInvoice", StringComparison.OrdinalIgnoreCase))
			   .OrderBy(x => x.SNo)
			   .ToList();
			//_logger.LogInformation($"Portal Data Count: {portalData.Count}");
			var summary = GenerateSLSummary(data);
			decimal[] grandTotal = getGrandTotal(data);

			var matchTypes = _configuration["SLMatchTypes"];
			var matchTypeList = matchTypes.Split(',').Select(x => x.Trim()).ToList();
			int count = matchTypeList.Count;


			using (var workbook = new XLWorkbook())
			{
				// ✅ Sheet 1: Summary
				var summarySheet = workbook.Worksheets.Add("Summary");
				//summarySheet.Range("A1:D1").Merge().Value = "Merged Cell Value";

				summarySheet.Range("D2:G2").Merge().Value = "SUM";
				summarySheet.Range("H2:K2").Merge().Value = "COUNT";

				//summarySheet.Range("D2:G2").Merge().Value = "Total Tax_A";
				//summarySheet.Range("H2:K2").Merge().Value = "Total Tax_A";


				summarySheet.Cell(3, 3).Value = "Data Source";
				//summarySheet.Cell(4, 1).Value = "Matching Results";
				summarySheet.Cell(4, 2).Value = "Categories";
				summarySheet.Cell(4, 3).Value = "Match Type";

				summarySheet.Range("D3:D4").Merge().Value = "Invoice";
				summarySheet.Range("E3:E4").Merge().Value = "EWayBill";
				summarySheet.Range("F3:F4").Merge().Value = "EInvoice";
				summarySheet.Range("G3:G4").Merge().Value = "Grand Total";

				summarySheet.Range("H3:H4").Merge().Value = "Invoice";
				summarySheet.Range("I3:I4").Merge().Value = "EWayBill";
				summarySheet.Range("J3:J4").Merge().Value = "EInvoice";
				summarySheet.Range("K3:K4").Merge().Value = "Grand Total";

				summarySheet.Range("L3:L4").Merge().Value = "% Matching";


				//summarySheet.Range("A1:B9").Merge();
				//summarySheet.Range("c1:c2").Merge();

				summarySheet.Range("A1:J4").Style.Font.Bold = true;
				summarySheet.Range("A1:J4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


				decimal totalInvoiceTax = Enumerable.Range(1, count).Sum(i => summary.GetType().GetProperty($"catagory{i}InvoiceSum")?.GetValue(summary) is decimal val ? val : 0);
				decimal totalEWayBilltax = Enumerable.Range(1, count).Sum(i => summary.GetType().GetProperty($"catagory{i}EWayWillSum")?.GetValue(summary) is decimal val ? val : 0);
				decimal totalEInvoicetax = Enumerable.Range(1, count).Sum(i => summary.GetType().GetProperty($"catagory{i}EInvoiceSum")?.GetValue(summary) is decimal val ? val : 0);

				decimal totalinvoicecount = Enumerable.Range(1, count).Sum(i => summary.GetType().GetProperty($"catagory{i}InvoiceNumber")?.GetValue(summary) is decimal val ? val : 0);
				decimal totalEWayBillcount = Enumerable.Range(1, count).Sum(i => summary.GetType().GetProperty($"catagory{i}EWayWillNumber")?.GetValue(summary) is decimal val ? val : 0);
				decimal totalEInvoicecount = Enumerable.Range(1, count).Sum(i => summary.GetType().GetProperty($"catagory{i}EInvoiceNumber")?.GetValue(summary) is decimal val ? val : 0);


				decimal grandtotaltax = 0;
				foreach (decimal value in grandTotal)
				{
					grandtotaltax += value;
				}
				decimal grandtotalcount = (decimal)totalinvoicecount + (decimal)totalEWayBillcount + (decimal)totalEInvoicecount;

				int summaryrow = 5;

				for (int i = 1; i <= count; i++)
				{
					var invoiceSum = GetDecimal(summary, $"catagory{i}InvoiceSum");
					var ewayWillSum = GetDecimal(summary, $"catagory{i}EWayWillSum");
					var einvoiceSum = GetDecimal(summary, $"catagory{i}EInvoiceSum");

					var invoiceNumber = GetDecimal(summary, $"catagory{i}InvoiceNumber");
					var ewayWillNumber = GetDecimal(summary, $"catagory{i}EWayWillNumber");
					var einvoiceNumber = GetDecimal(summary, $"catagory{i}EInvoiceNumber");

					AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[i - 1]), matchTypeList[i - 1], invoiceSum, ewayWillSum,
						einvoiceSum, invoiceNumber, ewayWillNumber, einvoiceNumber, grandtotalcount, grandTotal[i - 1]);
				}
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[0]), matchTypeList[0], summary.catagory1InvoiceSum ?? 0, summary.catagory1EWayWillSum ?? 0, summary.catagory1EInvoiceSum ?? 0 , summary.catagory1InvoiceNumber ?? 0, summary.catagory1EWayWillNumber ?? 0, summary.catagory1EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[1]), matchTypeList[1], summary.catagory2InvoiceSum ?? 0, summary.catagory2EWayWillSum ?? 0, summary.catagory2EInvoiceSum ?? 0 , summary.catagory2InvoiceNumber ?? 0, summary.catagory2EWayWillNumber ?? 0, summary.catagory2EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[2]), matchTypeList[2], summary.catagory3InvoiceSum ?? 0, summary.catagory3EWayWillSum ?? 0, summary.catagory3EInvoiceSum ?? 0 , summary.catagory3InvoiceNumber ?? 0, summary.catagory3EWayWillNumber ?? 0, summary.catagory3EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[3]), matchTypeList[3], summary.catagory4InvoiceSum ?? 0, summary.catagory4EWayWillSum ?? 0, summary.catagory4EInvoiceSum ?? 0 , summary.catagory4InvoiceNumber ?? 0, summary.catagory4EWayWillNumber ?? 0, summary.catagory4EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[4]), matchTypeList[4], summary.catagory5InvoiceSum ?? 0, summary.catagory5EWayWillSum ?? 0, summary.catagory5EInvoiceSum ?? 0 , summary.catagory5InvoiceNumber ?? 0, summary.catagory5EWayWillNumber ?? 0, summary.catagory5EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[5]), matchTypeList[5], summary.catagory6InvoiceSum ?? 0, summary.catagory6EWayWillSum ?? 0, summary.catagory6EInvoiceSum ?? 0 , summary.catagory6InvoiceNumber ?? 0, summary.catagory6EWayWillNumber ?? 0, summary.catagory6EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[6]), matchTypeList[6], summary.catagory7InvoiceSum ?? 0, summary.catagory7EWayWillSum ?? 0, summary.catagory7EInvoiceSum ?? 0 , summary.catagory7InvoiceNumber ?? 0, summary.catagory7EWayWillNumber ?? 0, summary.catagory7EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[7]), matchTypeList[7], summary.catagory8InvoiceSum ?? 0, summary.catagory8EWayWillSum ?? 0, summary.catagory8EInvoiceSum ?? 0 , summary.catagory8InvoiceNumber ?? 0, summary.catagory8EWayWillNumber ?? 0, summary.catagory8EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[8]), matchTypeList[8], summary.catagory9InvoiceSum ?? 0, summary.catagory9EWayWillSum ?? 0, summary.catagory9EInvoiceSum ?? 0 , summary.catagory9InvoiceNumber ?? 0, summary.catagory9EWayWillNumber ?? 0, summary.catagory9EInvoiceNumber ?? 0, grandtotalcount);


				summarySheet.Range($"A{summaryrow}:C{summaryrow}").Merge().Value = "Grand Total";
				summarySheet.Range($"A{summaryrow}:C{summaryrow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
				summarySheet.Cell(summaryrow, 4).Value = totalInvoiceTax;
				summarySheet.Cell(summaryrow, 5).Value = 1 * totalEWayBilltax;
				summarySheet.Cell(summaryrow, 6).Value = 1 * totalEInvoicetax;
				summarySheet.Cell(summaryrow, 7).Value = grandtotaltax;
				summarySheet.Cell(summaryrow, 8).Value = totalinvoicecount;
				summarySheet.Cell(summaryrow, 9).Value = totalEWayBillcount;
				summarySheet.Cell(summaryrow, 10).Value = totalEInvoicecount;
				summarySheet.Cell(summaryrow, 11).Value = grandtotalcount;
				//summarySheet.Cell(summaryrow, 12).Value = 0;

				//summarySheet.Range($"A1:J{summaryrow}").Style.Border.RightBorder = XLBorderStyleValues.Thin;
				//summarySheet.Range("A5:J13").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
				//summarySheet.Range($"A1:J{summaryrow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
				//summarySheet.Range($"A{summaryrow}:J{summaryrow}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;

				summaryrow = summaryrow + 3;
				summarySheet.Cell(summaryrow, 2).Value = "Request Created Date Time";
				summarySheet.Cell(summaryrow, 3).Value = "Request Updated Date Time";
				summarySheet.Cell(summaryrow, 4).Value = "Request Completed Date Time";

				summaryrow++;
				summarySheet.Cell(summaryrow, 2).Value = ticketDetails.RequestCreatedDate;
				summarySheet.Cell(summaryrow, 3).Value = ticketDetails.RequestUpdatedDate;
				summarySheet.Cell(summaryrow, 4).Value = ticketDetails.RequestCompleteDate;

				summarySheet.Cell(summaryrow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
				summarySheet.Cell(summaryrow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
				summarySheet.Cell(summaryrow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

				// ✅ Sheet 2: Main Output
				var mainSheet = workbook.Worksheets.Add("Main Output");
				var headers = new[]
				{
				 "Sno",
				 "User GSTIN",
				 "Compare",
				 "Datasource",
				 "Categories",
				 "Match Type",
				 "SupplierGSTIN",
				 "SupplierName",
				 "InvoiceNo",
				 "InvoiceDate",
				 "TaxableValue",
				 "CGST",
				 "SGST",
				 "IGST",
				 "CESS",
				 "TotalTax",
			 };
				// Write header row
				for (int i = 0; i < headers.Length; i++)
				{
					mainSheet.Cell(1, i + 1).Value = headers[i];
					mainSheet.Cell(1, i + 1).Style.Font.Bold = true;
				}
				int row = 2;
				// ✅ 1. Write Invoice Data (Sno starts from 1)
				foreach (var item in invoiceData)
				{
					int i = 1;
					mainSheet.Cell(row, i++).Value = item.SNo;
					mainSheet.Cell(row, i++).Value = item.ClientGstin;
					mainSheet.Cell(row, i++).Value = item.Compare;
					mainSheet.Cell(row, i++).Value = item.DataSource;
					mainSheet.Cell(row, i++).Value = item.Category;
					mainSheet.Cell(row, i++).Value = item.MatchType;
					mainSheet.Cell(row, i++).Value = item.CustomerGstin;
					mainSheet.Cell(row, i++).Value = item.CustomerName;
					mainSheet.Cell(row, i++).Value = item.InvoiceNumber;
					mainSheet.Cell(row, i++).Value = item.InvoiceDate;
					mainSheet.Cell(row, i++).Value = item.TaxableAmount;
					mainSheet.Cell(row, i++).Value = item.CGST;
					mainSheet.Cell(row, i++).Value = item.SGST;
					mainSheet.Cell(row, i++).Value = item.IGST;
					mainSheet.Cell(row, i++).Value = item.CESS;
					mainSheet.Cell(row, i++).Value = item.TotalAmount;

					row++;
				}

				// Write eWayBillData 
				foreach (var item in eWayBillData)
				{
					int i = 1;
					mainSheet.Cell(row, i++).Value = item.SNo;
					mainSheet.Cell(row, i++).Value = item.ClientGstin;
					mainSheet.Cell(row, i++).Value = item.Compare;
					mainSheet.Cell(row, i++).Value = item.DataSource;
					mainSheet.Cell(row, i++).Value = item.Category;
					mainSheet.Cell(row, i++).Value = item.MatchType;
					mainSheet.Cell(row, i++).Value = item.CustomerGstin;
					mainSheet.Cell(row, i++).Value = item.CustomerName;
					mainSheet.Cell(row, i++).Value = item.InvoiceNumber;
					mainSheet.Cell(row, i++).Value = item.InvoiceDate;
					mainSheet.Cell(row, i++).Value = item.TaxableAmount;
					mainSheet.Cell(row, i++).Value = item.CGST;
					mainSheet.Cell(row, i++).Value = item.SGST;
					mainSheet.Cell(row, i++).Value = item.IGST;
					mainSheet.Cell(row, i++).Value = item.CESS;
					mainSheet.Cell(row, i++).Value = item.TotalAmount;

					row++;
				}
				//eInvoiceData
				foreach (var item in eInvoiceData)
				{
					int i = 1;
					mainSheet.Cell(row, i++).Value = item.SNo;
					mainSheet.Cell(row, i++).Value = item.ClientGstin;
					mainSheet.Cell(row, i++).Value = item.Compare;
					mainSheet.Cell(row, i++).Value = item.DataSource;
					mainSheet.Cell(row, i++).Value = item.Category;
					mainSheet.Cell(row, i++).Value = item.MatchType;
					mainSheet.Cell(row, i++).Value = item.CustomerGstin;
					mainSheet.Cell(row, i++).Value = item.CustomerName;
					mainSheet.Cell(row, i++).Value = item.InvoiceNumber;
					mainSheet.Cell(row, i++).Value = item.InvoiceDate;
					mainSheet.Cell(row, i++).Value = item.TaxableAmount;
					mainSheet.Cell(row, i++).Value = item.CGST;
					mainSheet.Cell(row, i++).Value = item.SGST;
					mainSheet.Cell(row, i++).Value = item.IGST;
					mainSheet.Cell(row, i++).Value = item.CESS;
					mainSheet.Cell(row, i++).Value = item.TotalAmount;

					row++;
				}

				// Export Excel
				using (var stream = new MemoryStream())
				{
					workbook.SaveAs(stream);
					var content = stream.ToArray();

					//ViewBag.Message = "yes";

					return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
											$"{ExportFileName}");
				}
			}
			// return View("~/Views/Admin/CompareGstFiles/CompareGST.cshtml"); // Redirect to the same view if needed  
		}
		public async Task<Attachment> GenerateSLExcelAttachmentAsync(string requestNo, string fileName, string ClientGSTIN)
		{
			var ticketDetails = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			string ExportFileName = $"{fileName.Split('.')[0]}_{requestNo}_Report.xlsx";

			var data = await _sLComparedDataBusiness.GetComparedDataBasedOnTicketAsync(requestNo, ClientGSTIN);

			var invoiceData = data
				.Where(d => d.DataSource != null && d.DataSource.Trim().Equals("Invoice", StringComparison.OrdinalIgnoreCase))
				.OrderBy(x => x.SNo)
				.ToList();
			//_logger.LogInformation($"Invoice Data Count: {invoiceData.Count}");

			var eWayBillData = data
				.Where(d => d.DataSource != null && d.DataSource.Trim().Equals("EWayBill", StringComparison.OrdinalIgnoreCase))
				.OrderBy(x => x.SNo)
				.ToList();

			var eInvoiceData = data
			   .Where(d => d.DataSource != null && d.DataSource.Trim().Equals("EInvoice", StringComparison.OrdinalIgnoreCase))
			   .OrderBy(x => x.SNo)
			   .ToList();
			//_logger.LogInformation($"Portal Data Count: {portalData.Count}");
			var summary = GenerateSLSummary(data);
			decimal[] grandTotal = getGrandTotal(data);
			var matchTypes = _configuration["SLMatchTypes"];
			var matchTypeList = matchTypes.Split(',').Select(x => x.Trim()).ToList();
			int count = matchTypeList.Count;


			using (var workbook = new XLWorkbook())
			{
				// ✅ Sheet 1: Summary
				var summarySheet = workbook.Worksheets.Add("Summary");
				//summarySheet.Range("A1:D1").Merge().Value = "Merged Cell Value";

				summarySheet.Range("D2:G2").Merge().Value = "SUM";
				summarySheet.Range("H2:K2").Merge().Value = "COUNT";

				//summarySheet.Range("D2:G2").Merge().Value = "Total Tax_A";
				//summarySheet.Range("H2:K2").Merge().Value = "Total Tax_A";


				summarySheet.Cell(3, 3).Value = "Data Source";
				//summarySheet.Cell(4, 1).Value = "Matching Results";
				summarySheet.Cell(4, 2).Value = "Categories";
				summarySheet.Cell(4, 3).Value = "Match Type";

				summarySheet.Range("D3:D4").Merge().Value = "Invoice";
				summarySheet.Range("E3:E4").Merge().Value = "EWayBill";
				summarySheet.Range("F3:F4").Merge().Value = "EInvoice";
				summarySheet.Range("G3:G4").Merge().Value = "Grand Total";

				summarySheet.Range("H3:H4").Merge().Value = "Invoice";
				summarySheet.Range("I3:I4").Merge().Value = "EWayBill";
				summarySheet.Range("J3:J4").Merge().Value = "EInvoice";
				summarySheet.Range("K3:K4").Merge().Value = "Grand Total";

				summarySheet.Range("L3:L4").Merge().Value = "% Matching";


				//summarySheet.Range("A1:B9").Merge();
				//summarySheet.Range("c1:c2").Merge();

				summarySheet.Range("A1:J4").Style.Font.Bold = true;
				summarySheet.Range("A1:J4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


				decimal totalInvoiceTax = Enumerable.Range(1, count).Sum(i => summary.GetType().GetProperty($"catagory{i}InvoiceSum")?.GetValue(summary) is decimal val ? val : 0);
				decimal totalEWayBilltax = Enumerable.Range(1, count).Sum(i => summary.GetType().GetProperty($"catagory{i}EWayWillSum")?.GetValue(summary) is decimal val ? val : 0);
				decimal totalEInvoicetax = Enumerable.Range(1, count).Sum(i => summary.GetType().GetProperty($"catagory{i}EInvoiceSum")?.GetValue(summary) is decimal val ? val : 0);

				decimal totalinvoicecount = Enumerable.Range(1, count).Sum(i => summary.GetType().GetProperty($"catagory{i}InvoiceNumber")?.GetValue(summary) is decimal val ? val : 0);
				decimal totalEWayBillcount = Enumerable.Range(1, count).Sum(i => summary.GetType().GetProperty($"catagory{i}EWayWillNumber")?.GetValue(summary) is decimal val ? val : 0);
				decimal totalEInvoicecount = Enumerable.Range(1, count).Sum(i => summary.GetType().GetProperty($"catagory{i}EInvoiceNumber")?.GetValue(summary) is decimal val ? val : 0);


				decimal grandtotaltax = 0;
				foreach (decimal value in grandTotal)
				{
					grandtotaltax += value;
				}
				decimal grandtotalcount = (decimal)totalinvoicecount + (decimal)totalEWayBillcount + (decimal)totalEInvoicecount;

				int summaryrow = 5;

				for (int i = 1; i <= count; i++)
				{
					var invoiceSum = GetDecimal(summary, $"catagory{i}InvoiceSum");
					var ewayWillSum = GetDecimal(summary, $"catagory{i}EWayWillSum");
					var einvoiceSum = GetDecimal(summary, $"catagory{i}EInvoiceSum");

					var invoiceNumber = GetDecimal(summary, $"catagory{i}InvoiceNumber");
					var ewayWillNumber = GetDecimal(summary, $"catagory{i}EWayWillNumber");
					var einvoiceNumber = GetDecimal(summary, $"catagory{i}EInvoiceNumber");

					AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[i - 1]), matchTypeList[i - 1], invoiceSum, ewayWillSum,
						einvoiceSum, invoiceNumber, ewayWillNumber, einvoiceNumber, grandtotalcount, grandTotal[i - 1]);
				}
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[0]), matchTypeList[0], summary.catagory1InvoiceSum ?? 0, summary.catagory1EWayWillSum ?? 0, summary.catagory1EInvoiceSum ?? 0 , summary.catagory1InvoiceNumber ?? 0, summary.catagory1EWayWillNumber ?? 0, summary.catagory1EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[1]), matchTypeList[1], summary.catagory2InvoiceSum ?? 0, summary.catagory2EWayWillSum ?? 0, summary.catagory2EInvoiceSum ?? 0 , summary.catagory2InvoiceNumber ?? 0, summary.catagory2EWayWillNumber ?? 0, summary.catagory2EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[2]), matchTypeList[2], summary.catagory3InvoiceSum ?? 0, summary.catagory3EWayWillSum ?? 0, summary.catagory3EInvoiceSum ?? 0 , summary.catagory3InvoiceNumber ?? 0, summary.catagory3EWayWillNumber ?? 0, summary.catagory3EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[3]), matchTypeList[3], summary.catagory4InvoiceSum ?? 0, summary.catagory4EWayWillSum ?? 0, summary.catagory4EInvoiceSum ?? 0 , summary.catagory4InvoiceNumber ?? 0, summary.catagory4EWayWillNumber ?? 0, summary.catagory4EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[4]), matchTypeList[4], summary.catagory5InvoiceSum ?? 0, summary.catagory5EWayWillSum ?? 0, summary.catagory5EInvoiceSum ?? 0 , summary.catagory5InvoiceNumber ?? 0, summary.catagory5EWayWillNumber ?? 0, summary.catagory5EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[5]), matchTypeList[5], summary.catagory6InvoiceSum ?? 0, summary.catagory6EWayWillSum ?? 0, summary.catagory6EInvoiceSum ?? 0 , summary.catagory6InvoiceNumber ?? 0, summary.catagory6EWayWillNumber ?? 0, summary.catagory6EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[6]), matchTypeList[6], summary.catagory7InvoiceSum ?? 0, summary.catagory7EWayWillSum ?? 0, summary.catagory7EInvoiceSum ?? 0 , summary.catagory7InvoiceNumber ?? 0, summary.catagory7EWayWillNumber ?? 0, summary.catagory7EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[7]), matchTypeList[7], summary.catagory8InvoiceSum ?? 0, summary.catagory8EWayWillSum ?? 0, summary.catagory8EInvoiceSum ?? 0 , summary.catagory8InvoiceNumber ?? 0, summary.catagory8EWayWillNumber ?? 0, summary.catagory8EInvoiceNumber ?? 0, grandtotalcount);
				//AddSLRow(summarySheet, summaryrow++, GetSLCategory(matchTypeList[8]), matchTypeList[8], summary.catagory9InvoiceSum ?? 0, summary.catagory9EWayWillSum ?? 0, summary.catagory9EInvoiceSum ?? 0 , summary.catagory9InvoiceNumber ?? 0, summary.catagory9EWayWillNumber ?? 0, summary.catagory9EInvoiceNumber ?? 0, grandtotalcount);


				summarySheet.Range($"A{summaryrow}:C{summaryrow}").Merge().Value = "Grand Total";
				summarySheet.Range($"A{summaryrow}:C{summaryrow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
				summarySheet.Cell(summaryrow, 4).Value = totalInvoiceTax;
				summarySheet.Cell(summaryrow, 5).Value = 1 * totalEWayBilltax;
				summarySheet.Cell(summaryrow, 6).Value = 1 * totalEInvoicetax;
				summarySheet.Cell(summaryrow, 7).Value = grandtotaltax;
				summarySheet.Cell(summaryrow, 8).Value = totalinvoicecount;
				summarySheet.Cell(summaryrow, 9).Value = totalEWayBillcount;
				summarySheet.Cell(summaryrow, 10).Value = totalEInvoicecount;
				summarySheet.Cell(summaryrow, 11).Value = grandtotalcount;
				summarySheet.Cell(summaryrow, 12).Value = 0;

				//summarySheet.Range($"A1:J{summaryrow}").Style.Border.RightBorder = XLBorderStyleValues.Thin;
				//summarySheet.Range("A5:J13").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
				//summarySheet.Range($"A1:J{summaryrow}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
				//summarySheet.Range($"A{summaryrow}:J{summaryrow}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;

				summaryrow = summaryrow + 3;
				summarySheet.Cell(summaryrow, 2).Value = "Request Created Date Time";
				summarySheet.Cell(summaryrow, 3).Value = "Request Updated Date Time";
				summarySheet.Cell(summaryrow, 4).Value = "Request Completed Date Time";

				summaryrow++;
				summarySheet.Cell(summaryrow, 2).Value = ticketDetails.RequestCreatedDate;
				summarySheet.Cell(summaryrow, 3).Value = ticketDetails.RequestUpdatedDate;
				summarySheet.Cell(summaryrow, 4).Value = ticketDetails.RequestCompleteDate;

				summarySheet.Cell(summaryrow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
				summarySheet.Cell(summaryrow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
				summarySheet.Cell(summaryrow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

				// ✅ Sheet 2: Main Output
				var mainSheet = workbook.Worksheets.Add("Main Output");
				var headers = new[]
				{
		   "Sno",
		   "User GSTIN",
		   "Compare",
		   "Datasource",
		   "Categories",
		   "Match Type",
		   "SupplierGSTIN",
		   "SupplierName",
		   "InvoiceNo",
		   "InvoiceDate",
		   "TaxableValue",
		   "CGST",
		   "SGST",
		   "IGST",
		   "CESS",
		   "TotalTax",
	   };
				// Write header row
				for (int i = 0; i < headers.Length; i++)
				{
					mainSheet.Cell(1, i + 1).Value = headers[i];
					mainSheet.Cell(1, i + 1).Style.Font.Bold = true;
				}
				int row = 2;
				// ✅ 1. Write Invoice Data (Sno starts from 1)
				foreach (var item in invoiceData)
				{
					int i = 1;
					mainSheet.Cell(row, i++).Value = item.SNo;
					mainSheet.Cell(row, i++).Value = item.ClientGstin;
					mainSheet.Cell(row, i++).Value = item.Compare;
					mainSheet.Cell(row, i++).Value = item.DataSource;
					mainSheet.Cell(row, i++).Value = item.Category;
					mainSheet.Cell(row, i++).Value = item.MatchType;
					mainSheet.Cell(row, i++).Value = item.CustomerGstin;
					mainSheet.Cell(row, i++).Value = item.CustomerName;
					mainSheet.Cell(row, i++).Value = item.InvoiceNumber;
					mainSheet.Cell(row, i++).Value = item.InvoiceDate;
					mainSheet.Cell(row, i++).Value = item.TaxableAmount;
					mainSheet.Cell(row, i++).Value = item.CGST;
					mainSheet.Cell(row, i++).Value = item.SGST;
					mainSheet.Cell(row, i++).Value = item.IGST;
					mainSheet.Cell(row, i++).Value = item.CESS;
					mainSheet.Cell(row, i++).Value = item.TotalAmount;

					row++;
				}

				// Write eWayBillData 
				foreach (var item in eWayBillData)
				{
					int i = 1;
					mainSheet.Cell(row, i++).Value = item.SNo;
					mainSheet.Cell(row, i++).Value = item.ClientGstin;
					mainSheet.Cell(row, i++).Value = item.Compare;
					mainSheet.Cell(row, i++).Value = item.DataSource;
					mainSheet.Cell(row, i++).Value = item.Category;
					mainSheet.Cell(row, i++).Value = item.MatchType;
					mainSheet.Cell(row, i++).Value = item.CustomerGstin;
					mainSheet.Cell(row, i++).Value = item.CustomerName;
					mainSheet.Cell(row, i++).Value = item.InvoiceNumber;
					mainSheet.Cell(row, i++).Value = item.InvoiceDate;
					mainSheet.Cell(row, i++).Value = item.TaxableAmount;
					mainSheet.Cell(row, i++).Value = item.CGST;
					mainSheet.Cell(row, i++).Value = item.SGST;
					mainSheet.Cell(row, i++).Value = item.IGST;
					mainSheet.Cell(row, i++).Value = item.CESS;
					mainSheet.Cell(row, i++).Value = item.TotalAmount;

					row++;
				}
				//eInvoiceData
				foreach (var item in eInvoiceData)
				{
					int i = 1;
					mainSheet.Cell(row, i++).Value = item.SNo;
					mainSheet.Cell(row, i++).Value = item.ClientGstin;
					mainSheet.Cell(row, i++).Value = item.Compare;
					mainSheet.Cell(row, i++).Value = item.DataSource;
					mainSheet.Cell(row, i++).Value = item.Category;
					mainSheet.Cell(row, i++).Value = item.MatchType;
					mainSheet.Cell(row, i++).Value = item.CustomerGstin;
					mainSheet.Cell(row, i++).Value = item.CustomerName;
					mainSheet.Cell(row, i++).Value = item.InvoiceNumber;
					mainSheet.Cell(row, i++).Value = item.InvoiceDate;
					mainSheet.Cell(row, i++).Value = item.TaxableAmount;
					mainSheet.Cell(row, i++).Value = item.CGST;
					mainSheet.Cell(row, i++).Value = item.SGST;
					mainSheet.Cell(row, i++).Value = item.IGST;
					mainSheet.Cell(row, i++).Value = item.CESS;
					mainSheet.Cell(row, i++).Value = item.TotalAmount;

					row++;
				}

				var stream = new MemoryStream();

				workbook.SaveAs(stream);
				stream.Position = 0; // Reset stream before use

				var attachment = new Attachment(stream, ExportFileName,
					"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

				return attachment;

			}
		}

		private decimal GetDecimal(object obj, string propertyName)
		{
			var prop = obj.GetType().GetProperty(propertyName);
			var value = prop?.GetValue(obj);
			return value is decimal d ? d : 0;
		}

		private string GetSLCategory(string Match_Type)
		{
			//Console.WriteLine($"GetSLCategory function : {Match_Type}");
			string[] matchTypeList = _configuration["SLMatchTypes"]
				.Split(',')
				.Select(x => x.Trim())
				.ToArray();

			// Define match groups using indexes
			var completelyMatchedTypes = new[] { matchTypeList[0], matchTypeList[1] };
			var partiallyMatchedTypes = new[] { matchTypeList[2], matchTypeList[3], matchTypeList[4] };
			var unMatchedTypes = new[] { matchTypeList[5], matchTypeList[6], matchTypeList[7],
									matchTypeList[8], matchTypeList[9], matchTypeList[10], matchTypeList[11] };

			if (completelyMatchedTypes.Contains(Match_Type))
				return "Completely_Matched";
			else if (partiallyMatchedTypes.Contains(Match_Type))
				return "Partially_Matched";
			else if (unMatchedTypes.Contains(Match_Type))
				return "UnMatched";
			else
				return "Unknown";
		}

		public decimal[] getGrandTotal(List<SLComparedDataModel> data)
		{
			string[] matchTypeList = _configuration["SLMatchTypes"]
			  .Split(',')
			  .Select(x => x.Trim())
			  .ToArray();
			var categoryMap = new Dictionary<string, int>
			  {
				 { matchTypeList[0] , 1},
				 { matchTypeList[1] , 2},
				 { matchTypeList[2] , 3},
				 { matchTypeList[3] , 4},
				 { matchTypeList[4] , 5},
				 { matchTypeList[5] , 6},
				 { matchTypeList[6] , 7},
				 { matchTypeList[7] , 8},
				 { matchTypeList[8] , 9},
				 { matchTypeList[9] , 10},
				 { matchTypeList[10], 11},
				 { matchTypeList[11], 12},
			  };
			decimal[] grandTotal = new decimal[matchTypeList.Length];


			string[] comparisons = _configuration["Compare"].Split(',').Select(x => x.Trim()).ToArray();

			foreach (var kvp in categoryMap)
			{
				var categoryName = kvp.Key;
				var categoryNumber = kvp.Value;
				int i = categoryNumber - 1;


				var categoryData = data.Where(d => d.MatchType?.Trim().Equals(categoryName, StringComparison.OrdinalIgnoreCase) == true);

				var compare1INV = categoryData.Where(d => d.Compare != null &&
														  d.Compare.Trim().Equals(comparisons[0], StringComparison.OrdinalIgnoreCase) &&
														  d.DataSource != null &&
														  d.DataSource.Trim().Equals("Invoice", StringComparison.OrdinalIgnoreCase)).Sum(x => (decimal?)x.TotalAmount);
				var compare1EI = categoryData.Where(d => d.Compare != null &&
														 d.Compare.Trim().Equals(comparisons[0], StringComparison.OrdinalIgnoreCase) &&
														 d.DataSource != null &&
														 d.DataSource.Trim().Equals("EInvoice", StringComparison.OrdinalIgnoreCase)).Sum(x => (decimal?)x.TotalAmount);

				var compare2EI = categoryData.Where(d => d.Compare != null &&
														 d.Compare.Trim().Equals(comparisons[1], StringComparison.OrdinalIgnoreCase) &&
														 d.DataSource != null &&
														 d.DataSource.Trim().Equals("EInvoice", StringComparison.OrdinalIgnoreCase)).Sum(x => (decimal?)x.TotalAmount);
				var compare2EWB = categoryData.Where(d => d.Compare != null &&
														 d.Compare.Trim().Equals(comparisons[1], StringComparison.OrdinalIgnoreCase) &&
														 d.DataSource != null &&
														 d.DataSource.Trim().Equals("EWayBill", StringComparison.OrdinalIgnoreCase)).Sum(x => (decimal?)x.TotalAmount);

				var compare3INV = categoryData.Where(d => d.Compare != null &&
														 d.Compare.Trim().Equals(comparisons[2], StringComparison.OrdinalIgnoreCase) &&
														 d.DataSource != null &&
														 d.DataSource.Trim().Equals("Invoice", StringComparison.OrdinalIgnoreCase)).Sum(x => (decimal?)x.TotalAmount);
				var compare3EWB = categoryData.Where(d => d.Compare != null &&
														 d.Compare.Trim().Equals(comparisons[2], StringComparison.OrdinalIgnoreCase) &&
														 d.DataSource != null &&
														 d.DataSource.Trim().Equals("EWayBill", StringComparison.OrdinalIgnoreCase)).Sum(x => (decimal?)x.TotalAmount);

				grandTotal[i] = ((decimal)(compare1INV - compare1EI + compare2EI - compare2EWB + compare3INV - compare3EWB));
			}
			return grandTotal;
		}

		public SLComparedDataModel GenerateSLSummary(List<SLComparedDataModel> data)
		{
			//_logger.LogInformation("Data count : {0}", data.Count);


			var summary = new SLComparedDataModel();

			string[] matchTypeList = _configuration["SLMatchTypes"]
			   .Split(',')
			   .Select(x => x.Trim())
			   .ToArray();


			// Dictionary of category names mapped to numbers (to match your model)
			var categoryMap = new Dictionary<string, int>
		{
		   { matchTypeList[0] , 1},
		   { matchTypeList[1] , 2},
		   { matchTypeList[2] , 3},
		   { matchTypeList[3] , 4},
		   { matchTypeList[4] , 5},
		   { matchTypeList[5] , 6},
		   { matchTypeList[6] , 7},
		   { matchTypeList[7] , 8},
		   { matchTypeList[8] , 9},
		   { matchTypeList[9] , 10},
		   { matchTypeList[10], 11},
		   { matchTypeList[11], 12},


		};
			//_logger.LogInformation("Distinct categories in data: " + string.Join(", ", data.Select(d => d.Category).Distinct()));

			foreach (var kvp in categoryMap)
			{
				var categoryName = kvp.Key;
				var categoryNumber = kvp.Value;

				//var categoryData = data.Where(d => d.MatchType == categoryName);
				var categoryData = data.Where(d => d.MatchType?.Trim().Equals(categoryName, StringComparison.OrdinalIgnoreCase) == true);
				//Console.WriteLine($"categoryData group count : {categoryData.Count()}");

				//_logger.LogInformation("Distinct DataSource in data: " + string.Join(", ", data.Select(d => d.DataSource).Distinct()));

				var invoiceGroup = categoryData.Where(d => d.DataSource != null && d.DataSource.Trim().Equals("Invoice", StringComparison.OrdinalIgnoreCase));

				//Console.WriteLine($"invoice group count : {invoiceGroup.Count()}");

				var eWayBillGroup = categoryData.Where(d => d.DataSource != null && d.DataSource.Trim().Equals("EWayBill", StringComparison.OrdinalIgnoreCase));

				var eInvoiceGroup = categoryData.Where(d => d.DataSource != null && d.DataSource.Trim().Equals("EInvoice", StringComparison.OrdinalIgnoreCase));


				typeof(SLComparedDataModel).GetProperty($"catagory{categoryNumber}InvoiceNumber")?.SetValue(summary, (decimal)invoiceGroup.Count());
				typeof(SLComparedDataModel).GetProperty($"catagory{categoryNumber}InvoiceSum")?.SetValue(summary, invoiceGroup.Sum(x => (decimal?)x.TotalAmount) ?? 0);

				typeof(SLComparedDataModel).GetProperty($"catagory{categoryNumber}EWayWillNumber")?.SetValue(summary, (decimal)eWayBillGroup.Count());
				typeof(SLComparedDataModel).GetProperty($"catagory{categoryNumber}EWayWillSum")?.SetValue(summary, eWayBillGroup.Sum(x => (decimal?)x.TotalAmount) ?? 0);

				typeof(SLComparedDataModel).GetProperty($"catagory{categoryNumber}EInvoiceNumber")?.SetValue(summary, (decimal)eInvoiceGroup.Count());
				typeof(SLComparedDataModel).GetProperty($"catagory{categoryNumber}EInvoiceSum")?.SetValue(summary, eInvoiceGroup.Sum(x => (decimal?)x.TotalAmount) ?? 0);


			}

			//Console.WriteLine($"Summary in function : {summary.catagory1InvoiceSum}");
			return summary;
		}

		private void AddSLRow(IXLWorksheet sheet, int row, string catagory, string matchType, decimal invoiceSum, decimal EWaybillSum, decimal EInvoiceSum,
			decimal invoicecount, decimal Ewaybillcount, decimal EInvoicecount, decimal grandtotalcount, decimal grandtotal)
		{
			sheet.Cell(row, 2).Value = catagory;
			sheet.Cell(row, 3).Value = matchType;

			sheet.Cell(row, 4).Value = invoiceSum;
			sheet.Cell(row, 5).Value = EWaybillSum;
			sheet.Cell(row, 6).Value = EInvoiceSum;
			sheet.Cell(row, 7).Value = grandtotal;

			sheet.Cell(row, 8).Value = invoicecount;
			sheet.Cell(row, 9).Value = Ewaybillcount;
			sheet.Cell(row, 10).Value = EInvoicecount;
			sheet.Cell(row, 11).Value = invoicecount + Ewaybillcount + EInvoicecount;

			// Console.WriteLine(grandtotalcount);
			sheet.Cell(row, 12).Value = $"{Math.Round(((invoicecount + Ewaybillcount + EInvoicecount) / (decimal)grandtotalcount) * 100, 2)}%";

		}

		#endregion

		#region Function to redirect to Sales Register CompareSLCSVResults
		public async Task<IActionResult> CompareSLCSVResults(string requestNo, string ClientGSTIN)
		{
			ViewBag.Messages = "Admin";
			//Console.WriteLine($"ticketno" + requestNo);
			var Ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			ViewBag.Ticket = Ticket;
			var matchTypes = _configuration["SLMatchTypes"];
			var matchTypeList = matchTypes.Split(',').Select(x => x.Trim()).ToList();
			ViewBag.MatchType = matchTypeList;
			var data = await _sLComparedDataBusiness.GetComparedDataBasedOnTicketAsync(requestNo, ClientGSTIN);
			ViewBag.ReportDataList = data;
			//_logger.LogInformation("Total Compared Data Rows: " + data.Count);

			// store number and sum in a model and save it in 
			ViewBag.Summary = GenerateSLSummary(data);
			ViewBag.GrandTotal = getGrandTotal(data);
			//var summary = ViewBag.Summary;
			//_logger.LogInformation($"invoice 1 count {summary.catagory1InvoiceNumber},portal 1 count {summary.catagory1PortalNumber}");

			return View("~/Views/Admin/CompareSalesLedgerGstFiles/CompareSalesLedger.cshtml");
		}

		#endregion

		#region Download Sales Register EInvoice Sample Portal File
		[HttpGet]
		public IActionResult DownloadEInvoiceSampleFileCSV()
		{
			string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "SRSamplePortalEInvoice.csv");

			if (!System.IO.File.Exists(filePath))
			{
				return NotFound("Sample file not found.");
			}

			var fileBytes = System.IO.File.ReadAllBytes(filePath);
			return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SRSamplePortalEInvoice.csv");
		}
		[HttpGet]
		public IActionResult DownloadEInvoiceSampleFileXLS()
		{
			string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "SRSamplePortalEInvoice.xls");

			if (!System.IO.File.Exists(filePath))
			{
				return NotFound("Sample file not found.");
			}

			var fileBytes = System.IO.File.ReadAllBytes(filePath);
			return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SRSamplePortalEInvoice.xls");
		}
		[HttpGet]
		public IActionResult DownloadEInvoiceSampleFileXLSX()
		{
			string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "SRSamplePortalEInvoice.xlsx");

			if (!System.IO.File.Exists(filePath))
			{
				return NotFound("Sample file not found.");
			}

			var fileBytes = System.IO.File.ReadAllBytes(filePath);
			return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SRSamplePortalEInvoice.xlsx");
		}
		#endregion

		#region Download Sales Register EWayBill Sample Portal File
		[HttpGet]
		public IActionResult DownloadEWayBillSampleFileCSV()
		{
			string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "SRSamplePortalEWayBill.csv");

			if (!System.IO.File.Exists(filePath))
			{
				return NotFound("Sample file not found.");
			}

			var fileBytes = System.IO.File.ReadAllBytes(filePath);
			return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SRSamplePortalEWayBill.csv");
		}
		[HttpGet]
		public IActionResult DownloadEWayBillSampleFileXLS()
		{
			string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "SRSamplePortalEWayBill.xls");

			if (!System.IO.File.Exists(filePath))
			{
				return NotFound("Sample file not found.");
			}

			var fileBytes = System.IO.File.ReadAllBytes(filePath);
			return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SRSamplePortalEWayBill.xls");
		}
		[HttpGet]
		public IActionResult DownloadEWayBillSampleFileXLSX()
		{
			string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "SRSamplePortalEWayBill.xlsx");

			if (!System.IO.File.Exists(filePath))
			{
				return NotFound("Sample file not found.");
			}

			var fileBytes = System.IO.File.ReadAllBytes(filePath);
			return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SRSamplePortalEWayBill.xlsx");
		}

		#endregion

		public IActionResult DownloadFromGSTIN()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act2"; // For active button styling
			return View("~/Views/Future/ComingSoon.cshtml");
		}

		public IActionResult UploadToDb()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act2"; // For active button styling
			return View("~/Views/Future/ComingSoon.cshtml");
		}

		public IActionResult Payments()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act2"; // For active button styling
			return View("~/Views/Future/ComingSoon.cshtml");
		}

		public IActionResult GSTR1A()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act3"; // For active button styling
			return View("~/Views/Future/ComingSoon.cshtml");
		}

		public IActionResult GSTR2A()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act3"; // For active button styling
			return View("~/Views/Future/ComingSoon.cshtml");
		}

		public IActionResult GSTR2B()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act3"; // For active button styling
			return View("~/Views/Future/ComingSoon.cshtml");
		}

		public IActionResult ClientStatusReport()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act4"; // For active button styling
			return View("~/Views/Future/ComingSoon.cshtml");
		}

		public IActionResult ClientWiseMismatchReport()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act4"; // For active button styling
			return View("~/Views/Future/ComingSoon.cshtml");
		}

		public IActionResult ClientPayments()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act4"; // For active button styling
			return View("~/Views/Future/ComingSoon.cshtml");
		}

		public IActionResult GSTFilingReport()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act4"; // For active button styling
			return View("~/Views/Future/ComingSoon.cshtml");
		}

		public IActionResult Notices()
		{
			ViewBag.Messages = "Admin";
			ViewData["ActiveAction"] = "act5"; // For active button styling
			return View("~/Views/Future/ComingSoon.cshtml");
		}

        #region UploadNotice
        public IActionResult UploadNotice(string req, string Edit)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act7"; // For active button styling
            ViewBag.Message = req;
            ViewBag.Edit = Edit;
            return View("~/Views/Admin/NoticeTracker/UploadNotice.cshtml");
        }

        [HttpPost]
        public async Task<IActionResult> UploadNoticeFile(string noticeTitle, string noticeDate, string priority, string description, IFormFile uploadedFile, string requestNo)
        {
            // parameters - clientgstin, notice title,notice datetime, notice description, priority, pdf file path, status
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act7"; // For active button styling
            ViewBag.Edit = string.IsNullOrEmpty(requestNo) ? "No" : "Yes";
            string Edit = ViewBag.Edit;

            DateTime? noticeDateTime = null;
            if (DateTime.TryParseExact(noticeDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime temp))
            {
                noticeDateTime = temp;
            }
            //Console.WriteLine(noticeDateTime);

            // valadite 
            // create notecerequest number
            // model - noreqnumber,clientgstin,createddatetime,notice title,notice datetime,priority,notice description,pdf file path, status
            // save model into db using noreqnumber
            // call api to save pdf file using noreqnumber
            string req = string.IsNullOrEmpty(requestNo) ? GenerateNotice() : requestNo;
            DateTime? CreatedDate = DateTime.Now;
            if (Edit == "Yes")
            {
                var clientGstin = MySession.Current.gstin;
                var details = await _noticeDataBusiness.GetNoticeDetails(req, clientGstin);
                CreatedDate = details.CreatedDatetime;
            }
            var data = new NoticeDataModel
            {
                RequestNumber = req, // Generate a unique request number
                ClientGstin = MySession.Current.gstin, // Use the current user's GSTIN
                ClientName = MySession.Current.UserName, // Use the current user's name
                CreatedDatetime = CreatedDate, // Set the current date and time
                UpdatedDateTime = DateTime.Now,
                NoticeTitle = noticeTitle, // Initialize with an empty string
                NoticeDatetime = noticeDateTime, // Set the current date and time
                Priority = priority, // Default priority
                NoticeDescription = description, // Initialize with an empty string        
                status = "Open", // Default status
                FileName = uploadedFile?.FileName
            };

            try
            {
                await _noticeDataBusiness.saveNotices(data);
            }
            catch (Exception ex)
            {
                // Handle the exception (e.g., log it, show an error message)
                ViewBag.ErrorMessage = "An error occurred while creating the notice: " + ex.Message;
                return View("~/Views/Admin/NoticeTracker/UploadNotice.cshtml", data);
            }

            try
            {
                using (var httpClient = new HttpClient())
                {
                    using (var content = new MultipartFormDataContent())
                    {
                        // Add form fields
                        content.Add(new StringContent("EN"), "LanguageId");
                        content.Add(new StringContent(MySession.Current.Email), "EmailId");
                        content.Add(new StringContent(MySession.Current.UserName), "UserName");
                        content.Add(new StringContent(req), "TicketNumber");
                        content.Add(new StringContent("1"), "FileOrder");
                        content.Add(new StringContent($"File1 - {uploadedFile.FileName}"), "FileDescription");
                        content.Add(new StringContent(""), "ReserveIN1");
                        content.Add(new StringContent(""), "ReserveIN2");
                        content.Add(new StringContent(""), "ReserveIN3");
                        content.Add(new StringContent(MySession.Current.UserName), "LoginUser");
                        content.Add(new StringContent("INS"), "ModeFlag");
                        content.Add(new StringContent(""), "OldFileName");

                        // Add file
                        var fileContent = new StreamContent(uploadedFile.OpenReadStream());
                        fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
                        content.Add(fileContent, "FileObject", uploadedFile.FileName);

                        string apiUrl = $"{_configuration["ApiSettings:BaseUrl"]}{_configuration["ApiSettings:SavePdf"]}";
                        var response = await httpClient.PostAsync(apiUrl, content);

                        if (response.IsSuccessStatusCode)
                        {
                            Console.WriteLine("File uploaded successfully.");
                            var jsonResponse = await response.Content.ReadAsStringAsync();
                            //var result = JsonSerializer.Deserialize<FileUploadResponse>(jsonResponse);
                            //return result;
                        }
                        else
                        {
                            var error = await response.Content.ReadAsStringAsync();
                            throw new Exception($"Upload failed: {response.StatusCode} - {error}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle the exception (e.g., log it, show an error message)
                ViewBag.ErrorMessage = "An error occurred while uploading the file: " + ex.Message;
                return View("~/Views/Admin/NoticeTracker/UploadNotice.cshtml", data);
            }

            string gstin = MySession.Current.gstin;

            Attachment file = await GenerateNoticeFileAttachment(req, gstin);

            string toEmail = MySession.Current.Email;

            string subjectTemplate = _configuration["NoticeMail:SubjectTemplate"];
            string subject = subjectTemplate;

            string bodyTemplate = _configuration["NoticeMail:BodyTemplate"];
            string body = bodyTemplate
                 .Replace("{UserName}", MySession.Current.UserName)
                 .Replace("{NoticeNo}", req)
                 .Replace("{NoticeDate}", noticeDateTime?.ToString("dd-MMM-yyyy"))
				 .Replace("{NoticeCreatedDateTime}", CreatedDate?.ToString("dd-MMM-yyyy hh:mm:ss tt"))
                 .Replace("{NoticeDescription}", description);

            var mainAdminEmails = await _userBusiness.GetMainAdminUsers();  // helper to get MainAdmin clients
            string[] ccEmails = mainAdminEmails
								.Select(x => x.Email)
								.ToArray();

            await SendNoticeEmail(toEmail, subject, body, ccEmails, file);

            return RedirectToAction("UploadNotice", "Admin", new { req, Edit });
            //return View("~/Views/Admin/NoticeTracker/UploadNotice.cshtml");
        }

        private string GenerateNotice()
        {
            return "REQ_NT_" + DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        public async Task<IActionResult> EditUploadNotice(string requestNo, string ClientGSTIN)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act7";
            var notice = await _noticeDataBusiness.GetNoticeDetails(requestNo, ClientGSTIN);
            ViewBag.noticeDetails = notice;
            ViewBag.FileEdit = true;
            ViewBag.requestNo = requestNo;
            return View("~/Views/Admin/NoticeTracker/UploadNotice.cshtml");
        }

        public async Task SendNoticeEmail(string toEmail, string subject, string body, string[] ccEmails, Attachment attachment1)
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


                mail.To.Add(toEmail.Trim());


                if (ccEmails != null)
                {
                    foreach (var cc in ccEmails)
                    {
                        if (!string.IsNullOrWhiteSpace(cc))
                            mail.CC.Add(cc.Trim());
                        //continue;
                    }
                }

                // Attachments
                if (attachment1 != null)
                {
                    mail.Attachments.Add(attachment1);
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

        #region ActiveNotices
        public async Task<IActionResult> ActiveNotices(DateTime? fromdate, DateTime? todate, string status, string priority, string message)
        {
            // Default to 2 days ago and today if no values are provided
            DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
            DateTime toDateTime = todate ?? DateTime.Now; // Default to today
            string statusValue = status ?? "All"; // Default to "All" if not provided
            string priorityValue = priority ?? "All"; // Default to "All" if not provided
            string gstin = MySession.Current.gstin;

            //var email = MySession.Current.Email; // Get the email from session
            //var clients = await _userBusiness.GetAdminClients(email);
            //string[] gstinArray = clients.Select(c => c.ClientGSTIN).ToArray();

            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act5"; // For active button styling
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;
            ViewBag.status = status;
            ViewBag.priority = priority;
            ViewBag.gstin = gstin;
            //foreach (var gst in gstinArray)
            //{
            //	Console.WriteLine(gst);
            //}
            //Console.ReadKey();
            //ViewBag.gstin = gstinArray;
            bool flag = false;
            ViewBag.flag = flag;

            ViewBag.isServer = _configuration["IsServer"] == "true" ? true : false;
            ViewBag.serverURl = _configuration["ServerUrl"];

			if(fromdate.HasValue && todate.HasValue && fromdate > todate)
			{
                ViewBag.Error = "From Date cannot be greater than To Date.";
                ViewBag.UnreadCount = 0;

                // ✅ Prevent null reference in view
                ViewBag.Notices = new List<NoticeDataModel>(); // Replace with actual model type
                return View("~/Views/Admin/NoticeTracker/ActiveNotices.cshtml");
            }

            //var notices = await _noticeDataBusiness.GetClientsActiveNotices(gstinArray, fromDateTime, toDateTime, statusValue, priorityValue);

            // create helper to get data from bd using clientgstin,parameters and get in model formate
            // model - noreqnumber,notice title,notice datetime,uploded datetime, priority,status , action
            // using helper we get list of models
            // send it through viewbag to view

            var notices = await _noticeDataBusiness.GetActiveNotices(gstin, fromDateTime, toDateTime, statusValue, priorityValue);
            ViewBag.notices = notices;

            //Get unread notices for the clients
            //string source = "Client";
            //var unreadnotices = await _noticeDataBusiness.GetClientsUnreadNotices(gstinArray, source);
            //ViewBag.UnreadCount = unreadnotices.Count;

            string source = "MainAdmin"; // Assuming the source is "Client" for unread notices
            var unreadnotices = await _noticeDataBusiness.GetUnreadNotices(gstin, source);
            ViewBag.UnreadCount = unreadnotices.Count;

            ViewBag.Message = message;

            return View("~/Views/Admin/NoticeTracker/ActiveNotices.cshtml");
        }

        public async Task<IActionResult> UnreadNotices(string ClientGSTIN)
        {
            //Console.WriteLine("HI");
            string gstin = ClientGSTIN;
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act5"; // For active button styling

			//string[] gstinArray = System.Text.Json.JsonSerializer.Deserialize<string[]>(ClientGSTIN);

			//foreach (var gstin in gstins)
			//{
			//	Console.WriteLine(gstin);
			//}

			string source = "MainAdmin";
            var unreadnotices = await _noticeDataBusiness.GetUnreadNotices(gstin, source);
            //var unreadnotices = await _noticeDataBusiness.GetClientsUnreadNotices(gstins, source);

            ViewBag.notices = unreadnotices;

            bool flag = true;
            ViewBag.flag = flag;

            ViewBag.isServer = _configuration["IsServer"] == "true" ? true : false;
            ViewBag.serverURl = _configuration["ServerUrl"];

            return View("~/Views/Admin/NoticeTracker/ActiveNotices.cshtml");
        }

        #endregion

        #region Chat Messages
        public async Task<IActionResult> ConversationChat(string requestNo, string ClientGSTIN, string fromDate, string toDate, string priority, string status)
        {
            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act5"; // For active button styling

            var notice = await _noticeDataBusiness.GetNoticeDetails(requestNo, ClientGSTIN);
            // based on notice request number and clientgstin get all messages in chat db
            var chat = await _noticeDataBusiness.GetChatData(requestNo, ClientGSTIN);

            ViewBag.chat = chat; // Send the chat data to the view
            ViewBag.Notice = notice;
            ViewBag.userName = MySession.Current.UserName;
            ViewBag.fromDate = fromDate;
            ViewBag.toDate = toDate;
            ViewBag.priority = priority;
            ViewBag.status = status;

            ViewBag.isServer = _configuration["IsServer"] == "true" ? true : false;
            ViewBag.serverURl = _configuration["ServerUrl"];

            return View("~/Views/Admin/NoticeTracker/ConversationChat.cshtml");
        }

        public async Task<IActionResult> SaveChat(string message, string requestNumber, string ClientGstin)
        {
            //Console.WriteLine(MySession.Current.UserName);
            var data = new NoticeChatModel
            {
                RequestNumber = requestNumber, // Use the provided request number
                ClientGstin = ClientGstin, // Use the provided GSTIN
                Source = "Admin", // Set source as "Admin"
                Name = MySession.Current.UserName, // Use the current client's name
                Message = message, // Use the provided message
            };

            // save chat message model in db using noticeRequestnumber and clientgstin
            await _noticeDataBusiness.savaChatData(data);

            var details = await _noticeDataBusiness.GetNoticeDetails(requestNumber, ClientGstin);
            if (details.status == "Open")
            {
                await _noticeDataBusiness.UpdateNoticeStatus(requestNumber, ClientGstin);
            }


            return RedirectToAction("ConversationChat", "Admin"); // Redirect to ConversationChat after saving the chat message
        }

        [HttpPost]
        public async Task<IActionResult> MarkAsRead(string requestNumber, string gstin)
        {
            //Console.WriteLine($"MarkAsRead called with ReqNo: {requestNumber}, gstin: {gstin}");
            //await _noticeDataBusiness.MarkClientMessagesAsRead(requestNumber, gstin);
            await _noticeDataBusiness.MarkMainAdminMessagesAsRead(requestNumber, gstin);
            return Ok();
        }

        #endregion

        #region ClosedNotices and Close Notice
        public async Task<IActionResult> CloseNotice(string requestNo, string ClientGSTIN)
        {
			//Console.WriteLine($"CloseNotice called with ReqNo: {ReqNo}, gstin: {gstin}");
			// close notice using request number and clientgstin
			string closedBy = MySession.Current.UserName;
            await _noticeDataBusiness.CloseNotice(requestNo, ClientGSTIN, closedBy);
            // 2. Notify all clients in that group via SignalR
            await _hubContext.Clients.Group(requestNo).SendAsync("NoticeClosed", requestNo);

            return RedirectToAction("ClosedNotices", "Admin"); // Redirect to ActiveNotices after closing the notice
        }

        public async Task<IActionResult> ClosedNotices(DateTime? fromdate, DateTime? todate, string priority)
        {
            // Default to 2 days ago and today if no values are provided
            DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
            DateTime toDateTime = todate ?? DateTime.Now; // Default to today
            string priorityValue = priority ?? "All"; // Default to "All" if not provided

            var email = MySession.Current.Email; // Get the email from session
            var clients = await _userBusiness.GetAdminClients(email);
            string[] gstinArray = clients.Select(c => c.ClientGSTIN).ToArray();

            ViewBag.Messages = "Admin";
            ViewData["ActiveAction"] = "act5"; // For active button styling
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;
            ViewBag.priority = priority;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";

                // ✅ Prevent null reference in view
                ViewBag.Notices = new List<NoticeDataModel>(); // Replace with actual model type

                return View("~/Views/Admin/NoticeTracker/ClosedNotices.cshtml");
            }

            var notices = await _noticeDataBusiness.GetClientsClosedNotices(gstinArray, fromDateTime, toDateTime, priorityValue);

            // create helper to get data from bd using clientgstin,parameters and get in model formate
            // model - noreqnumber,notice title,notice datetime,uploded datetime, priority,status , action
            // using helper we get list of models
            // send it through viewbag to view

            ViewBag.notices = notices;

            return View("~/Views/Admin/NoticeTracker/ClosedNotices.cshtml");
        }

        #endregion

        #region PDF Download
        public async Task<IActionResult> PDF(string req, string gstin, string status, string priority, DateTime? fromDate, DateTime? toDate)
        {

            var client = await _userBusiness.GetUserDetails(gstin);

            var httpClient = new HttpClient();

            // Set up request URL

            string apiUrl = $"{_configuration["ApiSettings:BaseUrl"]}{_configuration["ApiSettings:GetPdf"]}";

            // Create request body (JSON)
            var requestData = new
            {
                LanguageId = "EN",
                UserName = client.USERNAME,
                EmailId = client.Email,
                TicketNumber = req,
                FileOrder = "1",
                ReserveIN1 = "",
                ReserveIN2 = "",
                ReserveIN3 = ""
            };

            // Serialize the request body to JSON
            var content = new StringContent(System.Text.Json.JsonSerializer.Serialize(requestData), Encoding.UTF8, "application/json");

            // Make the POST request (Note: Your curl uses GET, but sends a body — technically incorrect. HTTP spec disallows body in GET.)
            try
            {
                var response = await httpClient.PostAsync(apiUrl, content);
                var responseData = await response.Content.ReadAsStringAsync();
                JToken token = JToken.Parse(responseData);

                if (token.Type == JTokenType.Array)
                {
                    // It's a JSON Array
                    var result = JsonConvert.DeserializeObject<JArray>(responseData);

                    var fileObject = result.FirstOrDefault();

                    if (fileObject["returnStatus"]?.ToString() != "Success")
                    {
                        throw new Exception("Filed to download file");
                    }

                    var fileName = $"{req}_{fileObject["fileNameOUT"]}";
                    var fileBase64 = fileObject["fileObjectOUT"]?.ToString();

                    var fileBytes = Convert.FromBase64String(fileBase64);

                    return File(fileBytes, "application/pdf", fileName);
                }
                else if (token.Type == JTokenType.Object)
                {
                    // It's a JSON Object
                    var result2 = JsonConvert.DeserializeObject<JObject>(responseData);

                    if (result2["errorSource"].ToString() == "Failure")
                    {
                        //"errorMessage": "Please upload your file(s) to proceed with editing the Task details",

                        throw new Exception(result2["errorMessage"].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Error while downloading PDF: " + ex.Message;
                Console.WriteLine($"msg : {ViewBag.Message}");
            }

            return BadRequest(ViewBag.Message);

            //         ViewBag.Messages = "Admin";
            //         ViewData["ActiveAction"] = "act5"; // For active button styling
            //         ViewBag.Fromdate = fromDate;
            //         ViewBag.ToDate = toDate;
            //         ViewBag.status = status;
            //         ViewBag.priority = priority;
            //         //foreach (var gst in gstinArray)
            //         //{
            //         //	Console.WriteLine(gst);
            //         //}
            //         //Console.ReadKey();
            //         ViewBag.gstin = gstin;
            //         bool flag = false;
            //         ViewBag.flag = flag;

            //         ViewBag.isServer = _configuration["IsServer"] == "true" ? true : false;
            //         ViewBag.serverURl = _configuration["ServerUrl"];

            //         Console.WriteLine($"msg : {ViewBag.Message}");
            //return RedirectToAction("ActiveNotices", "Admin", new { message = ViewBag.Message });

        }

        public async Task<Attachment> GenerateNoticeFileAttachment(string req, string gstin)
        {
            Attachment attachment = null!;

            var client = await _userBusiness.GetUserDetails(gstin);
            var httpClient = new HttpClient();

            string apiUrl = $"{_configuration["ApiSettings:BaseUrl"]}{_configuration["ApiSettings:GetPdf"]}";

            var requestData = new
            {
                LanguageId = "EN",
                UserName = client.USERNAME,
                EmailId = client.Email,
                TicketNumber = req,
                FileOrder = "1",
                ReserveIN1 = "",
                ReserveIN2 = "",
                ReserveIN3 = ""
            };

            var content = new StringContent(
                System.Text.Json.JsonSerializer.Serialize(requestData),
                Encoding.UTF8,
                "application/json"
            );

            var response = await httpClient.PostAsync(apiUrl, content);
            var responseData = await response.Content.ReadAsStringAsync();
            JToken token = JToken.Parse(responseData);

            if (token.Type == JTokenType.Array)
            {
                var result = JsonConvert.DeserializeObject<JArray>(responseData);
                var fileObject = result.FirstOrDefault();

                if (fileObject["returnStatus"]?.ToString() != "Success")
                    throw new Exception("Failed to download file");

                // Get file name and extension
                var fileName = $"{req}_{fileObject["fileNameOUT"]}";
                var fileBase64 = fileObject["fileObjectOUT"]?.ToString();

                // If file name does not contain extension, try detecting it
                if (!Path.HasExtension(fileName))
                {
                    // Try to guess from MIME or default to .pdf
                    fileName += ".pdf";
                }

                var fileBytes = Convert.FromBase64String(fileBase64);

                // Detect MIME type dynamically
                string mimeType = GetMimeTypeFromFileName(fileName);

                var memoryStream = new MemoryStream(fileBytes);
                memoryStream.Position = 0;

                attachment = new Attachment(memoryStream, fileName, mimeType);
            }

            return attachment;
        }

        // Helper to detect MIME type
        private string GetMimeTypeFromFileName(string fileName)
        {
            var ext = Path.GetExtension(fileName)?.ToLowerInvariant();
            return ext switch
            {
                ".pdf" => "application/pdf",
                ".jpg" => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".png" => "image/png",
                ".txt" => "text/plain",
                ".doc" => "application/msword",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".xls" => "application/vnd.ms-excel",
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                _ => "application/octet-stream" // default
            };
        }

        #endregion

    }
}