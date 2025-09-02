namespace CAF.GstMatching.Web.Models
{
    public class HomeModal
    {
        public class RequestModel
        {
            public string LanguageId { get; set; }
            public string UserName { get; set; }
            public string ScreenName { get; set; }
            public string ReserveIN1 { get; set; }
            public string ReserveIN2 { get; set; }
            public string ReserveIN3 { get; set; }
        }

        public class ResponseModel
        {
            public string LanguageId { get; set; }
            public string ScreenName { get; set; }
            public string ReturnStatus { get; set; }
            public string ReturnMessage { get; set; }
            public string ControlName { get; set; }
            public string ControlType { get; set; }
            public string ControlCaption { get; set; }
            public string ControlHint { get; set; }
            public string Remarks { get; set; }
            public string ReserveOUT1 { get; set; }
            public string ReserveOUT2 { get; set; }
            public string ReserveOUT3 { get; set; }
        }
    }
}