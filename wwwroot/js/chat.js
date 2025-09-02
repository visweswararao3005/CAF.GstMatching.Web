"use strict";

var connection = new signalR.HubConnectionBuilder().withUrl("/chatHub").build();

document.getElementById("sendButton").disabled = true;

const addMessageIcon = document.getElementById('addMessageIcon');
const messageInputContainer = document.getElementById('messageInputContainer');
const sendMessageButton = document.getElementById('sendButton');
const messageInput = document.getElementById('messageInput');
const messages = document.getElementById('messages');
const senderName = "Sender's Name";
async function loadMessages(Language, username, emailid, taskId) {
    try {
        const response = await fetch("https://qwikflow.in/TechieJoe/api/TaskConversation", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify({ "LanguageId": Language, "UserName": username, "EmailId": emailid, "TaskId": taskId }),
        });
        const data = await response.json();
        messages.innerHTML = "";
        if (data && data.length > 0) {
            const noMessagesMessage = document.querySelector('.no-messages');
            if (noMessagesMessage) {
                noMessagesMessage.remove();
            }
            data.forEach((message) => {
                displayMessage(message.conversationText, message.conversationDateTime);
            });
            scrollToBottom();
        } else {
            const noMessagesMessage = document.createElement('div');
            noMessagesMessage.classList.add('message-block', 'no-messages');
            noMessagesMessage.innerText = 'No messages available';
            if (!document.querySelector('.no-messages')) {
                messages.appendChild(noMessagesMessage);
            }

            scrollToBottom();
        }
    } catch (error) {
        console.error('Error fetching messages:', error);
    }
}

function displayMessage(conversationText, conversationDateTime) {
    const noMessagesMessage = document.querySelector('.no-messages');
    if (noMessagesMessage) {
        noMessagesMessage.remove();
    }

    const messageBlock = document.createElement('div');
    messageBlock.classList.add('message-block');
    const senderInfo = document.createElement('div');
    senderInfo.classList.add('sender-info');
    senderInfo.innerHTML = `<strong>${senderName}</strong> <span style="font-size: 12px; color: black;">${conversationDateTime}</span>`;

    const messageContent = document.createElement('p');
    messageContent.textContent = conversationText;

    messages.appendChild(senderInfo);
    messageBlock.appendChild(messageContent);

    messages.appendChild(messageBlock);
    scrollToBottom();
}

addMessageIcon.addEventListener('click', function () {
    messageInputContainer.style.display = 'block';
    messageInput.focus();
});

sendMessageButton.addEventListener('click', async function (event) {
    event.preventDefault();
    const messageText = messageInput.value.replace(/'/g, "''").trim();
    if (messageText) {
        const currentDate = new Date();
        const dateFormatted = formatDate(currentDate);

        const Language = "EN";
        var chatContainer = document.getElementById('messages');
        var username = chatContainer.getAttribute('data-username');
        var emailid = chatContainer.getAttribute('data-email');
        const taskId = chatContainer.getAttribute('data-id');

        await saveMessage(Language, username, emailid, taskId, messageText, dateFormatted);

        connection.invoke("SendMessage", senderName, messageText).catch(function (err) {
            return console.error(err.toString());
        });

        messageInput.value = '';
        messageInputContainer.style.display = 'none';
        addMessageIcon.style.display = 'flex';
    } else {
        alert('Please enter a message');
    }
});
function formatDate(date) {
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const hours = date.getHours();
    const minutes = date.getMinutes();
    const isPM = hours >= 12;

    const hour12 = hours % 12 || 12;
    const minuteFormatted = minutes < 10 ? '0' + minutes : minutes;
    const ampm = isPM ? 'PM' : 'AM';

    const formattedDate = `${date.getDate()}-${months[date.getMonth()]}-${date.getFullYear()} ${hour12}:${minuteFormatted} ${ampm}`;
    return formattedDate;
}
async function saveMessage(Language, username, emailid, taskId, conversationText, conversationDateTime) {
    try {
        const response = await fetch("https://qwikflow.in/TechieJoe/api/SaveTaskConversation", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify({
                "LanguageId": Language,
                "UserName": username,
                "EmailId": emailid,
                "TaskId": taskId,
                "ConversationText": conversationText,
                "ConversationDateTime": conversationDateTime
            }),
        });
        const data = await response.json();
        if (data.returnStatus === "SUCCESS") {
            console.log("Message saved successfully");
        } else {
            console.error('Failed to save message:', data.returnMessage);
        }
    } catch (error) {
        console.error('Error saving message:', error);
    }
}

connection.on("ReceiveMessage", function (senderName, message) {
    const currentDate = new Date();
    const dateFormatted = currentDate.toLocaleString();
    displayMessage(message, dateFormatted);
});

connection.start().catch(function (err) {
    return console.error('SignalR connection failed: ', err);
});

const Language = "EN";
var chatContainer = document.getElementById('messages');
var username = chatContainer.getAttribute('data-username');
var emailid = chatContainer.getAttribute('data-email');
const taskId = chatContainer.getAttribute('data-id');

loadMessages(Language, username, emailid, taskId);
scrollToBottom();
function scrollToBottom() {
    const messages = document.getElementById('messages');
    messages.scrollTop = messages.scrollHeight;
}