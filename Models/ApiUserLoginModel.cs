namespace CAF.GstMatching.Web.Models
{
    public class ApiUserLoginModelfailedModel
    {
        public string errorSource { get; set; }
        public string errorNumber { get; set; }
        public string errorMessage { get; set; }
        public string additionalInfo1 { get; set; }
        public string additionalInfo2 { get; set; }
        public string additionalInfo3 { get; set; }
    }

    public class Login
    {
        public string languageId { get; set; }
        public string emailId { get; set; }
        public string password { get; set; }
        public string firstName { get; set; }
        public string middleName { get; set; }
        public string lastName { get; set; }
        public string userName { get; set; }
        public string returnStatus { get; set; }
        public string returnMessage { get; set; }
        public string reserveOUT1 { get; set; }
        public string reserveOUT2 { get; set; }
        public string reserveOUT3 { get; set; }
    }
}