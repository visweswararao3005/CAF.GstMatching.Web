using System.ComponentModel.DataAnnotations;

namespace CAF.GstMatching.Web.Models
{
    public class ChatMessage
    {
        [Key]
        public string FromUserName { get; set; }

        public string FromUserID { get; set; }
        public string ToUserID { get; set; }
        public string TaskID { get; set; }
        public string Message { get; set; }
        public string Status { get; set; }
        public DateTime CreatedOn { get; set; }
        public DateTime UpdatedOn { get; set; }
        public DateTime ViewedOn { get; set; }
        public bool IsActive { get; set; }
    }
}