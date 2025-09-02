namespace CAF.GstMatching.Web.Models
{
    public class RazorpayOrderModel
    {
        public string OrderId { get; set; }
        public string Key { get; set; }
        public decimal Amount { get; set; }
        public string Currency { get; set; }


        public string UserName { get; set; }
        public string UserEmail { get; set; }
        public string UserPhone { get; set; }
        public string UserAddress { get; set; }

    }
}
