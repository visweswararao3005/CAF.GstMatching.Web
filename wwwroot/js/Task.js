function costEngg() {
    ViewBag.CompletedTask = "TaskDetails";
    TjCaptions("CostEngg");
    ViewData["ActiveAction"] = "OrderNew";
    window.location = "/Task/CostEngg";
}
function detReq() {
    ViewBag.CompletedTask = "TaskDetails";
    TjCaptions("Requirement");
    ViewData["ActiveAction"] = "OrderNew";
    window.location = "/Task/DetReq";
}
function productScreen() {
    ViewBag.CompletedTask = "TaskDetails";
    TjCaptions("ProductSpecificScreen");
    ViewData["ActiveAction"] = "OrderNew";
    window.location = "/Task/ProductScreen";
}

function NewTaskCreation() {
    ViewBag.CompletedTask = "TaskDetails";
    TjCaptions("ProductSpecificScreen");
    ViewData["ActiveAction"] = "OrderNew";
    window.location = "/Task/NewTaskCreation";
}

function TjCaptions(string ScreenName) {
    string LanguageId = "EN";
    if (LanguageId != null && ScreenName != null) {
        var _httpClient = new HttpClient();
        var requestBody = new
            {
                LanguageId = LanguageId,
                UserName = HttpContext.Session.GetString("Email"),
                ScreenName = ScreenName
            };
                string jsonBody = JsonConvert.SerializeObject(requestBody);
        var baseUrl = "https://qwikflow.in/TechieJoe/api/Captions";
        _httpClient.BaseAddress = new Uri(baseUrl);
        var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
        try {
            var request = new HttpRequestMessage(HttpMethod.Get, baseUrl)
            {
                Content = content
            };
                    HttpResponseMessage response = _httpClient.Send(request);
            if (response.IsSuccessStatusCode) {
                var responseData = await response.Content.ReadAsStringAsync();
                var responseList = JsonConvert.DeserializeObject < List < ResponseModel >> (responseData);
                var responseDict = responseList.ToDictionary(item => item.ControlName, item => item.ControlCaption);
                ViewBag.ResponseDict = responseDict;
            }
            else {
                ViewBag.Message = "Invalid ScreenName";
                return View();
            }
        }
        catch (Exception ex)
        {
            ViewBag.Message = "Something missing";
            return View();
        }
    }

    return View();
}