using CAF.GstMatching.Web.Helpers;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Globalization;
using CAF.GstMatching.Business.Interface;
using CAF.GstMatching.Models.PurchaseTicketModel;
using CAF.GstMatching.Web.Common;
using System.Text;
using ClosedXML.Excel;
using DataTable = System.Data.DataTable;
using Path = System.IO.Path;
using System.Net.Http;
using DocumentFormat.OpenXml.Drawing.Charts;
using static System.Runtime.InteropServices.JavaScript.JSType;
using CAF.GstMatching.Models.CompareGst;
using CAF.GstMatching.Models;
using CAF.GstMatching.Business;
using Microsoft.AspNetCore.Authorization;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using CellType = NPOI.SS.UserModel.CellType;
using Org.BouncyCastle.Asn1.Ocsp;
using System.Net.Http.Headers;
using DocumentFormat.OpenXml.ExtendedProperties;
using Org.BouncyCastle.Ocsp;
using Newtonsoft.Json.Linq;
using Microsoft.AspNetCore.SignalR;
using CAF.GstMatching.Web.Hubs;
using Newtonsoft.Json;

namespace CAF.GstMatching.Web.Controllers
{
    [ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
    public class VendorController : Controller
    {
        private readonly ILogger<VendorController> _logger;
        private readonly HttpClient _httpClient;
        private readonly IHttpClientFactory _httpClientFactory;
        private readonly IConfiguration _configuration;
        private readonly IHubContext<ChatHub> _hubContext;

        private readonly IUserBusiness _userBusiness;

        private readonly IPurchaseDataBusiness _purchaseDataBusiness;
        private readonly IPurchaseTicketBusiness _purchaseTicketBusiness;
        private readonly ICompareGstBusiness _compareGstBusiness;
        private readonly IGSTR2DataBusiness _gSTR2DataBusiness;
        
        private readonly ISLDataBusiness _sLDataBusiness;
        private readonly ISLTicketsBusiness _sLTicketsBusiness;
        private readonly ISLEInvoiceBusiness _sLEInvoiceBusiness;
        private readonly ISLEWayBillBusiness _sLEWayBillBusiness;
        private readonly ISLComparedDataBusiness _sLComparedDataBusiness;

        private readonly INoticeDataBusiness _noticeDataBusiness;

       

        public VendorController(ILogger<VendorController> logger,
                                IHttpClientFactory httpClientFactory, 
                                IConfiguration configuration,
                                IHubContext<ChatHub> hubContext,

                                IUserBusiness userBusiness,

                                IPurchaseDataBusiness purchaseDataBusiness,
                                IPurchaseTicketBusiness purchaseTicketBusiness,
                                ICompareGstBusiness compareGstBusiness,
                                IGSTR2DataBusiness gSTR2DataBusiness,
                                
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
            _compareGstBusiness = compareGstBusiness;
            _gSTR2DataBusiness = gSTR2DataBusiness;

            _sLDataBusiness = sLDataBusiness;
            _sLTicketsBusiness = sLTicketsBusiness;
            _sLEInvoiceBusiness = sLEInvoiceBusiness;
            _sLEWayBillBusiness = sLEWayBillBusiness;
            _sLComparedDataBusiness = sLComparedDataBusiness;

            _noticeDataBusiness = noticeDataBusiness;
        }

        #region TjCaptions
        private async Task TjCaptions(string screenName)
        {
            var username = "na"; // Replace with actual username if needed
            var httpClient = _httpClientFactory.CreateClient();
            var captions = await CommonHelper.GetCaptionsAsync(screenName, username, _logger, _configuration, httpClient);
            ViewBag.ResponseDict = captions;
        }

        #endregion

        #region Dashboard
        // ✅ Vendor Dashboard
        //public IActionResult Dashboard()
        //{
        //	ViewBag.Messages = "Vendor";
        //	return View();
        //}

        public async Task<IActionResult> DashboardView()
        {
            ViewBag.Messages = "Vendor";
            string Period = DateTime.Now.AddMonths(-1).ToString("MMM-yy"); // Replace with actual period value
            var PIEData = await _compareGstBusiness.GetDashboardDataAsync(MySession.Current.gstin, Period);
            //Console.WriteLine($"total invoice count: {PIEData.TotalRecords}");
            //Console.WriteLine($"Completely_Matched count: {PIEData.MatchedRecordsCount}  , Amount : {PIEData.MatchedRecordsSum}");
            //Console.WriteLine($"Partially_Matched count: {PIEData.PartiallyMatchedRecordsCount} , Amount : {PIEData.PartiallyMatchedRecordsSum}");
            //Console.WriteLine($"Unmatched count: {PIEData.UnmatchedRecordsCount} , Amount : {PIEData.UnmatchedRecordsSum}");
            ViewBag.MRC = PIEData.MatchedRecordsCount;
            ViewBag.PRC = PIEData.PartiallyMatchedRecordsCount;
            ViewBag.UMC = PIEData.UnmatchedRecordsCount;
            ViewBag.MRS = PIEData.MatchedRecordsSum;
            ViewBag.PRS = PIEData.PartiallyMatchedRecordsSum;
            ViewBag.UMS = PIEData.UnmatchedRecordsSum;

            string toMonth = DateTime.Now.AddMonths(-1).ToString("MMM-yy"); // Replace with actual period value
            string fromMonth = DateTime.Now.AddMonths(-12).ToString("MMM-yy"); // Replace with actual period value
            string type = "paid"; // Replace with actual month value
            var BARData = await _compareGstBusiness.GetDashboardBARDataAsync(fromMonth, toMonth, MySession.Current.gstin, type);
            ViewBag.BARData = BARData;
            //Console.WriteLine($"BARData : {BARData}");
            //foreach (var item in BARData)
            //{
            //	Console.WriteLine($"Month: {item.Key}, Total Tax: {item.Value}");
            //}
            return View("Views/Vendor/Dashboard/Dashboard.cshtml"); // or however you're using it
        }

        public async Task<IActionResult> DashboardCharts(
	        string period,
	        string fromPeriod,
	        string toPeriod,
	        string monthShort,
	        string year,
	        string returnType,
	        string fromYear,
	        string fromMonth,
	        string toYear,
	        string toMonth,
	        string chartType)
		{
			ViewBag.Messages = "Vendor";

            //Console.WriteLine($"chartType: {chartType}");
            //chartType: paid
            //chartType: purchase
            //chartType: sales
            //
            //

			// Use monthShort and year directly; fallback to defaults if null
			string selectedMonth = monthShort ?? DateTime.Now.AddMonths(-1).ToString("MMM");
			string selectedYear = year ?? DateTime.Now.Year.ToString();

			ViewBag.SelectedMonth = selectedMonth;
			ViewBag.SelectedYear = selectedYear;
			ViewBag.ReturnType = returnType ?? "GSTR2A";
			ViewBag.FromYear = fromYear;
			ViewBag.FromMonth = fromMonth;
			ViewBag.ToYear = toYear;
			ViewBag.ToMonth = toMonth;
			ViewBag.ChartType = chartType;

            // PIE Chart Data
            var PIEData = new DashboardDataModel();

			if (chartType == "paid")
            {
				ViewBag.PIEErrorMessage = "No data available for the selected Type.";
				ViewBag.BarErrorMessage = "No data available for the selected Type.";
				return View("~/Views/Vendor/Dashboard/Dashboard.cshtml");
			}
            if (chartType == "purchase")
            {
				Console.WriteLine($"chartType : {chartType}");
				PIEData = await _compareGstBusiness.GetDashboardDataAsync(MySession.Current.gstin, period);
			}
            if (chartType == "sales")
            {
                Console.WriteLine($"chartType : {chartType}");
				PIEData = await _sLComparedDataBusiness.GetDashboardDataAsync(MySession.Current.gstin, period);
                Console.WriteLine($"PIEData : {PIEData.MatchedRecordsCount} - {PIEData.PartiallyMatchedRecordsCount} - {PIEData.UnmatchedRecordsCount} - {PIEData.MatchedRecordsSum} - {PIEData.PartiallyMatchedRecordsSum} - {PIEData.UnmatchedRecordsSum}");
			}

			if (PIEData == null )
			{
				//Console.WriteLine("No bar chart data available.");
				ViewBag.PIEErrorMessage = "No data available for the selected period.";				
			}
			else
			{
				ViewBag.MRC = PIEData.MatchedRecordsCount;
				ViewBag.PRC = PIEData.PartiallyMatchedRecordsCount;
				ViewBag.UMC = PIEData.UnmatchedRecordsCount;
				ViewBag.MRS = PIEData.MatchedRecordsSum;
				ViewBag.PRS = PIEData.PartiallyMatchedRecordsSum;
				ViewBag.UMS = PIEData.UnmatchedRecordsSum;
			}
			

			//Console.WriteLine($"Total invoice count: {PIEData.TotalRecords}");
			//Console.WriteLine($"Completely Matched count: {PIEData.MatchedRecordsCount}, Amount: {PIEData.MatchedRecordsSum}");
			//Console.WriteLine($"Partially Matched count: {PIEData.PartiallyMatchedRecordsCount}, Amount: {PIEData.PartiallyMatchedRecordsSum}");
			//Console.WriteLine($"Unmatched count: {PIEData.UnmatchedRecordsCount}, Amount: {PIEData.UnmatchedRecordsSum}");

			// Bar Chart Date Validation
			try
			{
				string fromDateStr = $"{fromMonth}-01-{fromYear}";
				string toDateStr = $"{toMonth}-01-{toYear}";

				DateTime fromDate = fromDate = DateTime.ParseExact(fromDateStr, "MM-dd-yyyy", null);
				DateTime toDate = DateTime.ParseExact(toDateStr, "MM-dd-yyyy", null);

				if (fromDate > toDate)
				{
					ViewBag.BarErrorMessage = "From Period cannot be greater than To Period";
					ViewBag.BARData = new Dictionary<string, decimal>(); // Initialize to avoid null
					return View("~/Views/Vendor/Dashboard/Dashboard.cshtml");
				}

				int monthDiff = (toDate.Year - fromDate.Year) * 12 + (toDate.Month - fromDate.Month);
				if (monthDiff >= 12)
				{
					ViewBag.BarErrorMessage = "From period and To period cannot be more than 12 months";
					ViewBag.BARData = new Dictionary<string, decimal>(); // Initialize to avoid null
					return View("~/Views/Vendor/Dashboard/Dashboard.cshtml");
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Date parsing error: {ex.Message}");
				ViewBag.BarErrorMessage = "Invalid date range provided.";
				ViewBag.BARData = new Dictionary<string, decimal>(); // Initialize to avoid null
				return View("~/Views/Vendor/Dashboard/Dashboard.cshtml");
			}

			var BARData = new Dictionary<string, decimal>();
			// BAR Chart Data
			//if (chartType == "paid")
			//{
			//	ViewBag.BarErrorMessage = "No data available for the selected .";
			//	ViewBag.BARData = new Dictionary<string, decimal>(); // Initialize to avoid null
			//	return View("~/Views/Vendor/Dashboard/Dashboard.cshtml");
			//}
			if (chartType == "purchase")
			{
				BARData = await _compareGstBusiness.GetDashboardBARDataAsync(fromPeriod, toPeriod, MySession.Current.gstin, returnType);
			}
			if (chartType == "sales")
			{
				BARData = await _sLComparedDataBusiness.GetDashboardBARDataAsync(fromPeriod, toPeriod, MySession.Current.gstin, returnType);
			}

			if (BARData == null || !BARData.Any())
			{
				//Console.WriteLine("No bar chart data available.");
				ViewBag.BarErrorMessage = "No data available for the selected period.";
				ViewBag.BARData = new Dictionary<string, decimal>(); // Initialize to avoid null
			}
			else
			{
				ViewBag.BARData = BARData;
			}

			return View("~/Views/Vendor/Dashboard/Dashboard.cshtml");
		}

        //public async Task<IActionResult> DashboardBARChart(
        //          string fromPeriod,
        //          string toPeriod,
        //          string monthShort,
        //          string year,
        //          string returnType,
        //          string fromYear,
        //          string fromMonth,
        //          string toYear,
        //          string toMonth,
        //          string chartType)
        //      {
        //	ViewBag.Messages = "Vendor";
        //          string selectedMonth = monthShort;
        //          string selectedYear = year;
        //          ViewBag.SelectedMonth = selectedMonth ;
        //          ViewBag.SelectedYear = selectedYear ;
        //          ViewBag.ReturnType = returnType ;
        //          ViewBag.FromYear = fromYear ;
        //          ViewBag.FromMonth = fromMonth ;
        //          ViewBag.ToYear = toYear ;
        //          ViewBag.ToMonth = toMonth ;
        //          ViewBag.ChartType = chartType ;

        //          Console.WriteLine("selectedMonth: " + selectedMonth);
        //          Console.WriteLine("selectedYear: " + selectedYear);
        //          Console.WriteLine("fromPeriod: " + fromPeriod);
        //          Console.WriteLine("toPeriod: " + toPeriod);
        //          Console.WriteLine("returnType: " + returnType);
        //          Console.WriteLine("fromYear: " + fromYear);
        //          Console.WriteLine("fromMonth: " + fromMonth);
        //          Console.WriteLine("toYear: " + toYear);
        //          Console.WriteLine("toMonth: " + toMonth); 

        //          var BARData = await _compareGstBusiness.GetDashboardBARDataAsync(fromPeriod, toPeriod, MySession.Current.gstin, returnType);
        //	ViewBag.BARData = BARData;

        //	return View("Views/Vendor/Dashboard/Dashboard.cshtml");
        //}

        #endregion

        #region Purchase Register Upload Invoice File
        // ✅ Compare GST Page
        public async Task<IActionResult> EditUploadGST(string requestNo)
        {
            var clientGSTIN = MySession.Current.gstin;
            var ticketDetails = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(requestNo, clientGSTIN);
            ViewBag.RequestNo = requestNo;
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act3";
            ViewBag.FinancialYear = ticketDetails.FinancialYear;
            ViewBag.PeriodType = ticketDetails.PeriodType;
            ViewBag.TxnPeriod = ticketDetails.TxnPeriod;
            ViewBag.FileName = ticketDetails.FileName;
            return View("~/Views/Vendor/PurchaseRegister/UploadGST.cshtml");
        }
        public IActionResult Upload(string ticketId, string Edit)
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act3";
            ViewBag.Message = ticketId;
            ViewBag.Edit = Edit;
            return View("~/Views/Vendor/PurchaseRegister/UploadGST.cshtml");
        }
        [HttpPost]
        public async Task<IActionResult> UploadGST(IFormFile gstFile, string financialYear, string periodtype, string period, string requestNo)
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act3";
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
                return View("~/Views/Vendor/PurchaseRegister/UploadGST.cshtml");
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
                        if (!ValidateColumnNames(dataTable))
                        {
                            ViewBag.ErrorMessage = "Please check columns names";
                            return View("~/Views/Vendor/PurchaseRegister/UploadGST.cshtml");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        return View("~/Views/Vendor/PurchaseRegister/UploadGST.cshtml");
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
                        if (!ValidateColumnNames(dataTable))
                        {
                            ViewBag.ErrorMessage = "Please check columns names";
                            return View("~/Views/Vendor/PurchaseRegister/UploadGST.cshtml");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        return View("~/Views/Vendor/PurchaseRegister/UploadGST.cshtml");
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
                        if (!ValidateColumnNames(dataTable))
                        {
                            ViewBag.ErrorMessage = "Please check columns names";
                            return View("~/Views/Vendor/PurchaseRegister/UploadGST.cshtml");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        return View("~/Views/Vendor/PurchaseRegister/UploadGST.cshtml");
                    }
                }
            }
            else
            {
                ViewBag.ErrorMessage = "Invalid file format. Please upload CSV/Xlsx/Xls file";
                return View("~/Views/Vendor/PurchaseRegister/UploadGST.cshtml");
            }

            //_logger.LogInformation("rows in datatable: {0}", dataTable.Rows.Count); 
            // Generate a ticket

            string name = MySession.Current.UserName;
            string usergstin = MySession.Current.gstin;
            ViewBag.ticketId = string.IsNullOrEmpty(requestNo) ? GenerateTicket() : requestNo;
            string ticketId = ViewBag.ticketId;
            try
            {
                await _purchaseDataBusiness.SavePurchaseDataAsync(dataTable, ticketId , usergstin);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"{ex.Message} ";
                return View("~/Views/Vendor/PurchaseRegister/UploadGST.cshtml");
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

            return RedirectToAction("Upload", "Vendor", new { ticketId, Edit });
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
        private bool ValidateColumnNames(DataTable dataTable)
        {
            var _invoice = _configuration["Invoice"];
            var _invoiceColumns = _invoice.Split(',').Select(x => x.Trim()).ToList();
            var uploadedColumns = dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName.ToLower()).ToList();
            return _invoiceColumns.All(col => uploadedColumns.Contains(col.ToLower()));
        }

        #endregion
        
        #region Purchase Register Current Requests
        public async Task<IActionResult> OpenTickets(DateTime? fromdate, DateTime? todate)
        {
           
            // Default to 2 days ago and today if no values are provided
            DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
            DateTime toDateTime = todate ?? DateTime.Now; // Default to today

            //await TjCaptions("OpenTickets"); // Load captions for OpenTask page
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act3"; // For active button styling

            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                return View("~/Views/Vendor/PurchaseRegister/OpenTickets.cshtml");
            }

            //ViewBag.UserGSTIN = HttpContext.Session.GetString("UserGSTIN");
            var Tickets = await _purchaseTicketBusiness.GetOpenTicketsStatusAsync(MySession.Current.gstin, fromDateTime, toDateTime); //(ViewBag.UserGSTIN); // Fetch pending tickets based on user_gstin

            ViewBag.Tickets = Tickets;

            return View("~/Views/Vendor/PurchaseRegister/OpenTickets.cshtml");
        }

        #endregion

        #region Purchase Register Closed Requests
        // ✅ Completed Tasks Page
        public async Task<IActionResult> ClosedRequests(DateTime? fromdate, DateTime? todate)
        {
            // Default to 2 days ago and today if no values are provided
            DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
            DateTime toDateTime = todate ?? DateTime.Now; // Default to today

            //await TjCaptions("CompletedTasks"); // Load captions for CompletedTask page
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act3"; // For active button styling
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                return View("~/Views/Vendor/PurchaseRegister/CompletedTickets.cshtml");
            }
            //ViewBag.UserGSTIN = HttpContext.Session.GetString("UserGSTIN");
            var Tickets = await _purchaseTicketBusiness.GetCloseTicketsStatusAsync(MySession.Current.gstin, fromDateTime, toDateTime); //(ViewBag.UserGSTIN); // Fetch pending tickets based on user_gstin

            ViewBag.Tickets = Tickets;
            
            return View("~/Views/Vendor/PurchaseRegister/CompletedTickets.cshtml");
        }

        #endregion

        #region Function to redirect to CompareCSVResults for vendor
        public async Task<IActionResult> CompareResults(string requestNo, string ClientGSTIN)
        {
            ViewBag.Messages = "Vendor";
            Console.WriteLine($"ticketno" + requestNo);
            // var ClientGSTIN = MySession.Current.gstin;
            var Ticket = await _purchaseTicketBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
            ViewBag.Ticket = Ticket;
            var clientGSTIN = Ticket.ClientGSTIN;
            var data = await _compareGstBusiness.GetComparedDataBasedOnTicketAsync(requestNo, clientGSTIN);
            ViewBag.ReportDataList = data;
            _logger.LogInformation("Total Compared Data Rows: " + data.Count);

            // store number and sum in a model and save it in 
            ViewBag.Summary = GenerateSummary(data);
            var summary = ViewBag.Summary;
            //_logger.LogInformation($"invoice 1 count {summary.catagory1InvoiceNumber},portal 1 count {summary.catagory1PortalNumber}");

            return View("/Views/Vendor/PurchaseRegister/CompareResults.cshtml");
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

        #endregion

        #region Download Purchase Register Sample Invoice File(CSV, XLS, XLSX)
        [HttpGet]
        public IActionResult DownloadSampleFileCSV()
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "PRSampleInvoice.csv");

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Sample file not found.");
            }

            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PRSampleInvoice.csv");
        }
        [HttpGet]
        public IActionResult DownloadSampleFileXLS()
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "PRSampleInvoice.xls");

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Sample file not found.");
            }

            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PRSampleInvoice.xls");
        }
        [HttpGet]
        public IActionResult DownloadSampleFileXLSX()
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "PRSampleInvoice.xlsx");

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Sample file not found.");
            }

            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PRSampleInvoice.xlsx");
        }
        #endregion
        
        #region ChangePasswordV
        public IActionResult ChangePasswordV()
        {
            ViewBag.Messages = "Vendor";
            return View("~/Views/Vendor/ChangePassword/ChangePasswordV.cshtml");
        }

        [HttpPost]
        public async Task<IActionResult> ChangePasswordV(string currentPassword, string newPassword, string confirmPassword)
        {
            ViewBag.Messages = "Vendor";
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
                    return View("~/Views/Vendor/ChangePassword/ChangePasswordV.cshtml");
                }
            }
            catch (Exception ex)
            {
                // Handle unexpected exceptions (e.g., network error, timeout)
                ViewBag.ErrorMessage = $"An error occurred while changing password: {ex.Message}";
                return View("~/Views/Vendor/ChangePassword/ChangePasswordV.cshtml");
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
                    return View("~/Views/Vendor/ChangePassword/ChangePasswordV.cshtml");
                }
                else
                {
                    // Handle failure (non-2xx response)
                    var responseObject = JsonConvert.DeserializeObject<dynamic>(responseData);
                    var errorMessage = responseObject?.errorMessage;
                    ViewBag.Message = errorMessage;
                    return View("~/Views/Vendor/ChangePassword/ChangePasswordV.cshtml");
                }
            }
            catch (Exception ex)
            {
                // Handle unexpected exceptions (e.g., network error, timeout)
                ViewBag.ErrorMessage = $"An error occurred while changing password: {ex.Message}";
                return View("~/Views/Vendor/ChangePassword/ChangePasswordV.cshtml");
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

        #region SalesLedgerUploadInvoiceFile
        public async Task<IActionResult> SalesLedgerUploadFile(string requestNo)
        {
            var clientGSTIN = MySession.Current.gstin;
            var ticketDetails = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(requestNo, clientGSTIN);
            ViewBag.RequestNo = requestNo;
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act2";
            ViewBag.FinancialYear = ticketDetails.FinancialYear;
            ViewBag.PeriodType = ticketDetails.PeriodType;
            ViewBag.TxnPeriod = ticketDetails.Period;
            ViewBag.FileName = ticketDetails.FileName;
            return View("~/Views/Vendor/SalesRegister/UploadSalesLedgerFile.cshtml");
        }

        public IActionResult SalesLedgerUpload(string ticketId, string Edit)
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act2";
            ViewBag.Message = ticketId;
            ViewBag.Edit = Edit;
            return View("~/Views/Vendor/SalesRegister/UploadSalesLedgerFile.cshtml");
        }

        [HttpPost]
        public async Task<IActionResult> SalesLedgerUploadFile(string financialYear, string periodtype, string period, string requestNo, IFormFile gstFile)
        {
            //
            // Console.WriteLine("SalesLedgerUploadFileAsync");
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act2"; // For active button styling
            ViewBag.requestNo = requestNo;
            ViewBag.Edit = string.IsNullOrEmpty(requestNo) ? "No" : "Yes";
            if (gstFile == null || gstFile.Length == 0)
            {
                ViewBag.ErrorMessage = "Please select a valid file.";
                return View("~/Views/Vendor/SalesRegister/UploadSalesLedgerFile.cshtml");
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
                            return View("~/Views/Vendor/SalesRegister/UploadSalesLedgerFile.cshtml");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        return View("~/Views/Vendor/SalesRegister/UploadSalesLedgerFile.cshtml");
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
                            return View("~/Views/Vendor/SalesRegister/UploadSalesLedgerFile.cshtml");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        return View("~/Views/Vendor/SalesRegister/UploadSalesLedgerFile.cshtml");
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
                            return View("~/Views/Vendor/SalesRegister/UploadSalesLedgerFile.cshtml");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        return View("~/Views/Vendor/SalesRegister/UploadSalesLedgerFile.cshtml");
                    }
                }
            }
            else
            {
                ViewBag.ErrorMessage = "Invalid file format. Please upload CSV/Xlsx/Xls file";
                return View("~/Views/Vendor/SalesRegister/UploadSalesLedgerFile.cshtml");
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
                return View("~/Views/Vendor/SalesRegister/UploadSalesLedgerFile.cshtml");
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
            return RedirectToAction("SalesLedgerUpload", "Vendor", new { ticketId, Edit });
            // return View("~/Views/Vendor/SalesRegister/UploadSalesLedgerFile.cshtml");
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
            foreach(var list in _SLinvoiceColumns)
            {
                Console.WriteLine(list);
            }

            var uploadedColumns = dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName.ToLower()).ToList();
            foreach (var col in _SLinvoiceColumns)
            {
                Console.WriteLine(col.ToLower());
            }
            return _SLinvoiceColumns.All(col => uploadedColumns.Contains(col.ToLower()));
        }

        #endregion

        #region SalesLedgerCurrentRequests
        public async Task<IActionResult> SalesLedgerCurrentRequests(DateTime? fromdate, DateTime? todate)
        {
            // Default to 2 days ago and today if no values are provided
            DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
            DateTime toDateTime = todate ?? DateTime.Now; // Default to today

            //await TjCaptions("SLCR"); // Load captions for OpenTask page
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act2"; // For active button styling
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                return View("~/Views/Vendor/SalesRegister/SalesLedgerCurrentRequests.cshtml");
            }
            var Tickets = await _sLTicketsBusiness.GetOpenTicketsStatusAsync(MySession.Current.gstin, fromDateTime, toDateTime); //(ViewBag.UserGSTIN); // Fetch pending tickets based on user_gstin

            //Console.WriteLine("SalesLedgerCureentRequests" + Tickets.Count);
            //Console.WriteLine("Request no " + Tickets[0].RequestNumber);

            ViewBag.Tickets = Tickets;           

            return View("~/Views/Vendor/SalesRegister/SalesLedgerCurrentRequests.cshtml");
        }

        #endregion

        #region SalesLedgerClosedRequests
        public async Task<IActionResult> SalesLedgerClosedRequests(DateTime? fromdate, DateTime? todate)
        {
            // Default to 2 days ago and today if no values are provided
            DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
            DateTime toDateTime = todate ?? DateTime.Now; // Default to today

            //await TjCaptions("CompletedTasks"); // Load captions for CompletedTask page
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act2"; // For active button styling

            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                return View("~/Views/Vendor/SalesRegister/SalesLedgerClosedRequests.cshtml");
            }
            var Tickets = await _sLTicketsBusiness.GetCloseTicketsStatusAsync(MySession.Current.gstin, fromDateTime, toDateTime); //(ViewBag.UserGSTIN); // Fetch pending tickets based on user_gstin
           // Console.WriteLine("Tickets" + Tickets[0].RequestNumber);
            ViewBag.Tickets = Tickets;

            return View("~/Views/Vendor/SalesRegister/SalesLedgerClosedRequests.cshtml");
        }

        #endregion

        #region Function to redirect to CompareSRResults
        public async Task<IActionResult> CompareSRResults(string requestNo, string ClientGSTIN)
        {
            ViewBag.Messages = "Vendor";
            var Ticket = await _sLTicketsBusiness.GetUserDataBasedOnTicketAsync(requestNo, ClientGSTIN);
            ViewBag.Ticket = Ticket;
            var matchTypes = _configuration["SLMatchTypes"];
            var matchTypeList = matchTypes.Split(',').Select(x => x.Trim()).ToList();
            ViewBag.MatchType = matchTypeList;
            
            // var clientGSTIN = Ticket.ClientGstin;
            var data = await _sLComparedDataBusiness.GetComparedDataBasedOnTicketAsync(requestNo, ClientGSTIN);
            ViewBag.ReportDataList = data;
            //_logger.LogInformation("Total Compared Data Rows: " + data.Count);

            // store number and sum in a model and save it in 
            ViewBag.Summary = GenerateSLSummary(data);
            ViewBag.GrandTotal = getGrandTotal(data);
            //var summary = ViewBag.Summary;
            //_logger.LogInformation($"invoice 1 count {summary.catagory1InvoiceNumber},portal 1 count {summary.catagory1PortalNumber}");

            return View("/Views/Vendor/SalesRegister/CompareSRResults.cshtml");
        }

        public SLComparedDataModel GenerateSLSummary(List<SLComparedDataModel> data)
        {
            //_logger.LogInformation("Data count : {0}", data.Count);


            var summary = new SLComparedDataModel();

            // Dictionary of category names mapped to numbers (to match your model)
            var categoryMap = new Dictionary<string, int>
         {
             { "1_Exactly_Matched_GST_INV_DT_TAXB_TAX", 1 },
             { "2_Matched_With_Tolerance_GST_INV_DT_TAXB_TAX", 2 },
             { "3_Partially_Matched_GST_INV", 3 },
             { "4_Partially_Matched_GST_DT", 4 },
             { "5_Probable_Matched_GST_TAXB", 5 },
             { "6_UnMatched_Excess_or_Short_1_Invoicewise", 6 },
             { "Available_In_Sales_Register_MISSING_In_E_Invoice", 7 },
             { "Available_In_Sales_Register_MISSING_In_E_WayBIll", 8 },
             { "Available_In_E_Invoice_MISSING_In_Sales_Register", 9 },
             { "Available_In_E_Invoice_MISSING_In_E_WayBIll", 10 },
             { "Available_In_E_WayBIll_MISSING_In_E_Invoice", 11 },
             { "Available_In_E_WayBIll_MISSING_In_Sales_Register", 12 },
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

        #endregion

        #region Download Sales Register Sample Invoice File(CSV, XLS, XLSX)
        [HttpGet]
        public IActionResult DownloadSRSampleFileCSV()
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "SRSampleInvoice.csv");

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Sample file not found.");
            }

            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SRSampleInvoice.csv");
        }
        [HttpGet]
        public IActionResult DownloadSRSampleFileXLS()
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "SRSampleInvoice.xls");

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Sample file not found.");
            }

            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SRSampleInvoice.xls");
        }
        [HttpGet]
        public IActionResult DownloadSRSampleFileXLSX()
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "SampleFiles", "SRSampleInvoice.xlsx");

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Sample file not found.");
            }

            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SRSampleInvoice.xlsx");
        }

        #endregion

        public IActionResult CNDN()
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act4"; // For active button styling
            return View("~/Views/Future/ComingSoon.cshtml");
        }

        public IActionResult Reconciliation()
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act5"; // For active button styling
            return View("~/Views/Future/ComingSoon.cshtml");
        }

        public IActionResult GSTR1A()
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act5"; // For active button styling
            return View("~/Views/Future/ComingSoon.cshtml");
        }

        public IActionResult GSTR2A()
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act5"; // For active button styling
            return View("~/Views/Future/ComingSoon.cshtml");
        }

        public IActionResult GSTR2B()
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act5"; // For active button styling
            return View("~/Views/Future/ComingSoon.cshtml");
        }

        public IActionResult ReportGSTR1A()
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act6"; // For active button styling
            return View("~/Views/Future/ComingSoon.cshtml");
        }

        public IActionResult ReportGSTR2A()
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act6"; // For active button styling
            return View("~/Views/Future/ComingSoon.cshtml");
        }

        public IActionResult ReportGSTR2B()
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act6"; // For active button styling
            return View("~/Views/Future/ComingSoon.cshtml");
        }

        public IActionResult ExceptionReports()
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act6"; // For active button styling
            return View("~/Views/Future/ComingSoon.cshtml");
        }

        #region UploadNotice
        public IActionResult UploadNotice(string req, string Edit)
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act7"; // For active button styling
            ViewBag.Message = req;
            ViewBag.Edit = Edit;
            return View("~/Views/Vendor/NoticeTracker/UploadNotice.cshtml");
        }

        [HttpPost]
        public async Task<IActionResult> UploadNoticeFile(string noticeTitle, string noticeDate, string priority, string description, IFormFile uploadedFile, string requestNo)
        {
            // parameters - clientgstin, notice title,notice datetime, notice description, priority, pdf file path, status
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act7"; // For active button styling
            ViewBag.Edit = string.IsNullOrEmpty(requestNo) ? "No" : "Yes";
            string Edit = ViewBag.Edit;

            DateTime? noticeDateTime = null;
            if (DateTime.TryParseExact(noticeDate, "yyyy-MM-dd", CultureInfo.InvariantCulture,DateTimeStyles.None, out DateTime temp))
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
                return View("~/Views/Vendor/NoticeTracker/UploadNotice.cshtml", data);
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
                return View("~/Views/Vendor/NoticeTracker/UploadNotice.cshtml", data);
            }

            return RedirectToAction("UploadNotice", "Vendor", new { req, Edit });
            //return View("~/Views/Vendor/NoticeTracker/UploadNotice.cshtml");
        }

        private string GenerateNotice()
        {
            return "REQ_NT_" + DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        public async Task<IActionResult> EditUploadNotice(string requestNo, string ClientGSTIN)
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act7";
            var notice = await _noticeDataBusiness.GetNoticeDetails(requestNo, ClientGSTIN);
            ViewBag.noticeDetails = notice;
            ViewBag.FileEdit = true;
            ViewBag.requestNo = requestNo;
            return View("~/Views/Vendor/NoticeTracker/UploadNotice.cshtml");
        }

        #endregion

        #region Delete Pdf File
        public async Task<IActionResult> DeletePdfFile(string requestNumber, string ClientGstin)
        {
            Console.WriteLine($"Request Number: {requestNumber}, Client GSTIN: {ClientGstin}");

            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act7"; // For active button styling

            var notice = await _noticeDataBusiness.GetNoticeDetails(requestNumber, ClientGstin);

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
                        content.Add(new StringContent(requestNumber), "TicketNumber");
                        content.Add(new StringContent("1"), "FileOrder");
                        content.Add(new StringContent($"Delete - {notice.FileName}"), "FileDescription");
                        content.Add(new StringContent(""), "ReserveIN1");
                        content.Add(new StringContent(""), "ReserveIN2");
                        content.Add(new StringContent(""), "ReserveIN3");
                        content.Add(new StringContent(MySession.Current.UserName), "LoginUser");
                        content.Add(new StringContent("DEL"), "ModeFlag");
                        content.Add(new StringContent(notice.FileName), "OldFileName");

                        //// Add file
                        //var fileContent = new StreamContent(uploadedFile.OpenReadStream());
                        //fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
                        //content.Add(fileContent, "FileObject", uploadedFile.FileName);

                        string apiUrl = $"{_configuration["ApiSettings:BaseUrl"]}{_configuration["ApiSettings:SavePdf"]}";
                        var response = await httpClient.PostAsync(apiUrl, content);
                        var jsonString = await response.Content.ReadAsStringAsync();
                        // Parse JSON string
                        var result = JObject.Parse(jsonString);

                        // Check return status
                        if (result["returnStatus"]?.ToString() == "SUCCESS")
                        {
                            await _noticeDataBusiness.DeleteFileNameInNotice(requestNumber, ClientGstin);
                            Console.WriteLine("Filename removed from database.");
                        }
                        else
                        {
                            Console.WriteLine("File deletion failed: " + result["errorMessage"]);
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                // Handle the exception (e.g., log it, show an error message)
                ViewBag.ErrorMessage = "An error occurred while uploading the file: " + ex.Message;
                return View();
            }


            // Redirect to the ActiveNotices page after deletion
            return RedirectToAction("EditUploadNotice", "Vendor", new { requestNo = requestNumber, ClientGSTIN = ClientGstin});
        }

        #endregion

        #region ActiveNotices
        public async Task<IActionResult> ActiveNotices(DateTime? fromdate, DateTime? todate, string status, string priority)
        {
            // Default to 2 days ago and today if no values are provided
            DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
            DateTime toDateTime = todate ?? DateTime.Now; // Default to today
            string statusValue = status ?? "All"; // Default to "All" if not provided
            string priorityValue = priority ?? "All"; // Default to "All" if not provided
            string gstin = MySession.Current.gstin; // Get the current user's GSTIN

            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act7"; // For active button styling     
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;
            ViewBag.status = status;
            ViewBag.priority = priority;
            ViewBag.Gstin = gstin;
            bool flag = false;
            ViewBag.flag = flag;

            ViewBag.isServer = _configuration["IsServer"] == "true" ? true : false;
            ViewBag.serverURl = _configuration["ServerUrl"];

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";
                ViewBag.UnreadCount = 0;

                // ✅ Prevent null reference in view
                ViewBag.Notices = new List<NoticeDataModel>(); // Replace with actual model type
               
                return View("~/Views/Vendor/NoticeTracker/ActiveNotices.cshtml");
            }

            var notices = await _noticeDataBusiness.GetActiveNotices(gstin, fromDateTime, toDateTime, statusValue, priorityValue);
            ViewBag.Notices = notices;

            //Get Unread notices
            string source = "Admin"; // Assuming the source is "Client" for unread notices
            var unreadnotices = await _noticeDataBusiness.GetUnreadNotices(gstin,source);
            ViewBag.UnreadCount = unreadnotices.Count;
          
            //Console.WriteLine("Active Notices Count: " + notices.Count);
            //Console.WriteLine("Unread Notices Count: " + unreadnotices.Count);

            // Just filter for unread
            //var unreadNotices = notices.Where(n => n.IsRead == false).ToList();
            
            return View("~/Views/Vendor/NoticeTracker/ActiveNotices.cshtml");
        }

        public async Task<IActionResult> UnreadNotices(string ClientGSTIN)
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act7"; // For active button styling

            string gstin = ClientGSTIN;
            string source = "Admin";
            var unreadnotices = await _noticeDataBusiness.GetUnreadNotices(gstin,source);
            
            //Console.WriteLine("Unread Notices for " + gstin + ": " + unreadnotices.Count);
            //Console.WriteLine("Unread Notices Count: " + unreadnotices.Count);

            ViewBag.Gstin = gstin;

            ViewBag.Notices = unreadnotices;

            bool flag = true;
            ViewBag.flag = flag;

            ViewBag.isServer = _configuration["IsServer"] == "true" ? true : false;
            ViewBag.serverURl = _configuration["ServerUrl"];

            return View("~/Views/Vendor/NoticeTracker/ActiveNotices.cshtml");

        }
        #endregion

        #region ChatMessages

        public async Task<IActionResult> ConversationChat(string requestNo, string ClientGSTIN, string fromDate, string toDate, string priority, string status)
        {
            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act7"; // For active button styling
           
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

            return View("~/Views/Vendor/NoticeTracker/ConversationChat.cshtml");
        }

        public async Task<IActionResult> SaveChat(string message, string requestNumber, string ClientGstin)
        {
            Console.WriteLine(MySession.Current.UserName);
            var data = new NoticeChatModel
            {
                RequestNumber = requestNumber, // Use the provided request number
                ClientGstin = ClientGstin, // Use the provided GSTIN
                Source = "Client", // Set source as "Client"
                Name = MySession.Current.UserName, // Use the current client's name
                Message = message, // Use the provided message
            };

            // save chat message model in db using noticeRequestnumber and clientgstin
            await _noticeDataBusiness.savaChatData(data);

            return RedirectToAction("ConversationChat", "Vendor"); // Redirect to ConversationChat after saving the chat message
        }

        [HttpPost]
        public async Task<IActionResult> MarkAsRead(string requestNumber, string gstin)
        {
            await _noticeDataBusiness.MarkAdminMessagesAsRead(requestNumber, gstin);
            return Ok();
        }

        #endregion

        #region ClosedNotices
        public async Task<IActionResult> CloseNotice(string requestNo, string ClientGSTIN)
        {
            //Console.WriteLine($"CloseNotice called with ReqNo: {ReqNo}, gstin: {gstin}");
            // close notice using request number and clientgstin
            string closedBy = MySession.Current.UserName;
            await _noticeDataBusiness.CloseNotice(requestNo, ClientGSTIN, closedBy);
            // 2. Notify all clients in that group via SignalR
            await _hubContext.Clients.Group(requestNo).SendAsync("NoticeClosed", requestNo);

            return RedirectToAction("ClosedNotices", "Vendor"); // Redirect to ActiveNotices after closing the notice
        }
        public async Task<IActionResult> ClosedNotices(DateTime? fromdate, DateTime? todate, string priority)
        {
            // Default to 2 days ago and today if no values are provided
            DateTime fromDateTime = fromdate ?? DateTime.Now.AddDays(-1); // Default to 2 days ago
            DateTime toDateTime = todate ?? DateTime.Now; // Default to today
            string priorityValue = priority ?? "All"; // Default to "All" if not provided
            string gstin = MySession.Current.gstin; // Get the current user's GSTIN

            ViewBag.Messages = "Vendor";
            ViewData["ActiveAction"] = "act7"; // For active button styling
            ViewBag.Fromdate = fromdate;
            ViewBag.ToDate = todate;
            ViewBag.priority = priority;

            if (fromdate.HasValue && todate.HasValue && fromdate > todate)
            {
                ViewBag.Error = "From Date cannot be greater than To Date.";

                // ✅ Prevent null reference in view
                ViewBag.Notices = new List<NoticeDataModel>(); // Replace with actual model type

                return View("~/Views/Vendor/NoticeTracker/ClosedNotices.cshtml");
            }

            var notices = await _noticeDataBusiness.GetClosedNotices(gstin, fromDateTime, toDateTime, priorityValue);

            ViewBag.Notices = notices;

            return View("~/Views/Vendor/NoticeTracker/ClosedNotices.cshtml");
        }

        #endregion

    }
}