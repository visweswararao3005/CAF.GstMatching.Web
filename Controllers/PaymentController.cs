using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.AspNetCore.Mvc;
using System.Security.Cryptography;
using System.Text;
using Razorpay.Api;
using System;
using CAF.GstMatching.Business.Interface;
using Order = Razorpay.Api.Order;
using CAF.GstMatching.Web.Models;
using CAF.GstMatching.Web.Common;
using CAF.GstMatching.Models.UserModel;
using CAF.GstMatching.Data;
using System.Threading.Tasks;


namespace CAF.GstMatching.Web.Controllers
{
    public class PaymentController : Controller
    {
        private readonly IConfiguration _configuration;
        //private readonly PaymentHelper _paymentHelper;

        private readonly IUserBusiness _userBusiness;

        public PaymentController(IConfiguration config,
                                 IUserBusiness userBusiness
            )
        {
            _configuration = config;
            //_paymentHelper = new PaymentHelper();
            _userBusiness = userBusiness;
        }

        public async Task<IActionResult> Index()
        {
            var user = await _userBusiness.getUserValidUpto(MySession.Current.Email);
            ViewBag.userValidUpto = user.accessToDate;
            if (ViewBag.userValidUpto >= DateTime.Today)
            {
                ViewBag.Messages = "Admin";
            }
            // You can return a view or any other response here
            ViewBag.PricePerMonth = int.Parse(_configuration["Payment:PricePerMonth"]);
            ViewBag.Plans = _configuration.GetSection("Payment:Plans").Get<List<int>>(); // e.g., [1,2,3,6,12]
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> CreateOrder(decimal amount,decimal months)
        {
            try
            {
                int daysPerMonth = int.Parse(_configuration["Payment:DaysPerMonth"]);
                int days = (int)(months * daysPerMonth); // Convert months to days

                var user = await _userBusiness.GetUserNameAsync(MySession.Current.Email);

                var keyId = _configuration["Razorpay:KeyId"];
                var keySecret = _configuration["Razorpay:KeySecret"];

                string TransactionId = "tnx_id_" + Guid.NewGuid().ToString("N").Substring(0, 10).ToUpper();

                RazorpayClient client = new RazorpayClient(keyId, keySecret);

                Dictionary<string, object> options = new Dictionary<string, object>();
                options.Add("amount", amount * 100); // Razorpay works in paise
                options.Add("currency", "INR");
                options.Add("receipt", TransactionId);
                options.Add("payment_capture", 1);

                Order order = client.Order.Create(options);


                // id,ClientName,ClientGstin,ClientEmail,ClientPhone,ClientAddress,Amount,days,orderId,PaymentId,PaymentStatus,PaymentDate,TransactionId
                var paymentDetails = new PaymentDetailModel
                {
                    ClientName = user.USERNAME,
                    ClientGstin = user.Designation,
                    ClientEmail = user.Email,
                    ClientPhone = user.MobilePIN,
                    ClientAddress = user.Address,
                    Amount = amount,
                    Days = days,
                    OrderId = order["id"].ToString(),
                    PaymentId = string.Empty, // Will be filled after payment verification
                    PaymentStatus = "Created",
                    PaymentDate = DateTime.Now,
                    TransactionId = TransactionId // Will be filled after payment verification
                };

               await _userBusiness.savePaymentDetails(paymentDetails);

                var orderModel = new RazorpayOrderModel
                {
                    OrderId = order["id"].ToString(),
                    Key = keyId,
                    Amount = amount,
                    Currency = "INR",

                    UserName = user.USERNAME,
                    UserEmail = user.Email,
                    UserPhone = user.MobilePIN,
                };

                return Json(orderModel);
            }
            catch (Exception ex)
            {
                return BadRequest(new { error = ex.Message });
            }
        }

        [HttpPost]
        public async Task<IActionResult> VerifyPayment([FromBody] RazorpayPaymentResponse payment)
        {
            try
            {
                string email = MySession.Current.Email;
                string gstin = MySession.Current.gstin;
                string orderId = payment.razorpay_order_id;
                
                var keySecret = _configuration["Razorpay:KeySecret"];
                string generatedSignature;

                using (var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(keySecret)))
                {
                    var data = payment.razorpay_order_id + "|" + payment.razorpay_payment_id;
                    var hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(data));
                    generatedSignature = BitConverter.ToString(hash).Replace("-", "").ToLower();
                }

                if (generatedSignature == payment.razorpay_signature)
                {
                    var paymentDetails = await _userBusiness.GetPaymentDetails(orderId,email,gstin);
                    var paymentData = new PaymentDetailModel
                    {
                        ClientName = paymentDetails.ClientName,
                        ClientGstin = paymentDetails.ClientGstin,
                        ClientEmail = paymentDetails.ClientEmail,
                        OrderId = orderId,

                        PaymentId = payment.razorpay_payment_id,
                        PaymentStatus = "Paid",
                        PaymentDate = DateTime.Now,
                        TransactionId = payment.razorpay_signature
                    };
                    await _userBusiness.savePaymentDetails(paymentData);

                    // Add days to his uservalideupto table
                    int days = (int)paymentDetails.Days;
                    await _userBusiness.UpdateUserValidUpto(email, days);

                    return Ok(new { status = "success" });
                }
                else
                {
                    var paymentDetails = await _userBusiness.GetPaymentDetails(orderId, email, gstin);
                    var paymentData = new PaymentDetailModel
                    {
                        ClientName = paymentDetails.ClientName,
                        ClientGstin = paymentDetails.ClientGstin,
                        ClientEmail = paymentDetails.ClientEmail,
                        OrderId = orderId,

                        PaymentId = payment.razorpay_payment_id,
                        PaymentStatus = "failed",
                        PaymentDate = DateTime.Now,
                        TransactionId = payment.razorpay_signature
                    };
                    await _userBusiness.savePaymentDetails(paymentData);
                    //_userBusiness.SavePaymentDetails(payment.razorpay_order_id, "Failed");
                    return BadRequest(new { status = "failed" });
                }
            }
            catch (Exception ex)
            {
                return BadRequest(new { error = ex.Message });
            }
        }

    }

    public class RazorpayPaymentResponse
    {
        public string razorpay_order_id { get; set; }
        public string razorpay_payment_id { get; set; }
        public string razorpay_signature { get; set; }
    }

}
