function ordernewTask() {
    TjCaptions("Ordernewtask");
    ViewData["ActiveAction"] = "OrderNew";
    window.location = "/Task/OrderNew";
}
function openTask() {
    ViewData["ActiveAction"] = "openTask";
    TjCaptions("OpenTasks");
    window.location = "/Task/openTask";
}
function completedTask() {
    ViewData["ActiveAction"] = "completedTask";
    TjCaptions("CompletedTasks");
    window.location = "/Task/CompletedTask";
}

function logOut() {
    alert("Are You Sure Want to Logout");
    window.location = "/Home/LogOut";
}