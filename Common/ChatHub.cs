using Microsoft.AspNetCore.SignalR;
using System.Threading.Tasks;

namespace CAF.GstMatching.Web.Hubs
{
    public class ChatHub : Hub
    {
        /// <summary>
        /// Called when a user opens a chat for a particular request number
        /// Adds user to a SignalR group named after the request number
        /// </summary>
        public async Task JoinGroup(string requestNumber)
        {
            await Groups.AddToGroupAsync(Context.ConnectionId, requestNumber);
        }

        /// <summary>
        /// Called when a user sends a message
        /// Sends message to all users in that request number group
        /// </summary>
        public async Task SendMessageToGroup(string requestNumber, string sender, string message)
        {
            
            var time = DateTime.Now.ToString("dd-MM-yyyy hh:mm tt");
            await Clients.Group(requestNumber).SendAsync("ReceiveMessage", requestNumber, sender, message, time);

        }


        /// <summary>
        /// Helper method to format time
        /// </summary>
        private string GetFormattedTime()
        {
            return System.DateTime.Now.ToString("hh:mm tt"); // e.g., 11:45 AM
        }

        /// <summary>
        /// Optional: Let user leave group if you want to implement that
        /// </summary>
        public async Task LeaveGroup(string requestNumber)
        {
            await Groups.RemoveFromGroupAsync(Context.ConnectionId, requestNumber);
        }
        /// <summary>
        /// Called when Admin closes the notice
        /// Broadcasts a message to the group that the notice is closed
        /// </summary>
        public async Task NotifyNoticeClosed(string requestNumber)
        {
            await Clients.Group(requestNumber).SendAsync("NoticeClosed", requestNumber);
        }

        /// <summary>
        /// to notify other users that the sender is typing
        /// </summary>  
        public async Task SendTypingNotification(string requestNumber, string sender)
        {
            await Clients.Group(requestNumber).SendAsync("UserTyping", requestNumber, sender);
        }
    }
}
