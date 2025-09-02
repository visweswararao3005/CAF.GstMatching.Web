using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Net;
using System.Text;
using static CAF.GstMatching.Web.Models.HomeModal;

namespace CAF.GstMatching.Web.Helpers
{
    public static class CommonHelper
    {
        public static async Task<CaptionsResult> GetCaptionsAsync(
           string screenName,
           string userName,
           ILogger logger,
           IConfiguration configuration,
           HttpClient httpClient,
           string languageId = "EN")
        {
            var requestBody = new
            {
                LanguageId = languageId,
                UserName = userName,
                ScreenName = screenName
            };

            string jsonBody = JsonConvert.SerializeObject(requestBody);
            string baseUrl = configuration["ApiSettings:BaseUrl"] ?? throw new InvalidOperationException("ApiSettings:BaseUrl is not configured");
            string captionsEndpoint = configuration["ApiSettings:CaptionsEndpoint"] ?? throw new InvalidOperationException("ApiSettings:CaptionsEndpoint is not configured");
            string apiUrl = $"{baseUrl}{captionsEndpoint}";
            var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            try
            {
                var response = await httpClient.PostAsync(apiUrl, content);
                var responseData = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    // Try to parse as List<ResponseModel>
                    try
                    {
                        var responseList = JsonConvert.DeserializeObject<List<ResponseModel>>(responseData);
                        if (responseList != null && responseList.Any())
                        {
                            var captionsDict = responseList.ToDictionary(item => item.ControlName, item => item.ControlCaption);
                            return new CaptionsResult
                            {
                                Captions = captionsDict,
                                StatusCode = response.StatusCode
                            };
                        }
                    }
                    catch
                    {
                        // fall through
                    }

                    return new CaptionsResult
                    {
                        Captions = null,
                        StatusCode = response.StatusCode,
                        ErrorMessage = "No captions found in Captions API response."
                    };
                }
                else
                {
                    // Attempt to parse error JSON
                    try
                    {
                        var errorObj = JsonConvert.DeserializeObject<ErrorResponseModel>(responseData);
                        return new CaptionsResult
                        {
                            Captions = null,
                            StatusCode = response.StatusCode,
                            ErrorMessage = errorObj?.errorMessage ?? $"Captions API Failed"
                        };
                    }
                    catch
                    {
                        return new CaptionsResult
                        {
                            Captions = null,
                            StatusCode = response.StatusCode,
                            ErrorMessage = $"Captions API {response.StatusCode}"
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                return new CaptionsResult
                {
                    Captions = null,
                    StatusCode = HttpStatusCode.InternalServerError,
                    ErrorMessage = ex.Message
                };
            }
        }


    }

    public class ResponseModel
    {
        public string ControlName { get; set; }
        public string ControlCaption { get; set; }
    }

    public class ErrorResponseModel
    {
        public string errorSource { get; set; }
        public string errorNumber { get; set; }
        public string errorMessage { get; set; }
        public string additionalInfo1 { get; set; }
        public string additionalInfo2 { get; set; }
        public string additionalInfo3 { get; set; }
    }

    public class CaptionsResult
    {
        public Dictionary<string, string> Captions { get; set; }
        public HttpStatusCode StatusCode { get; set; }
        public string ErrorMessage { get; set; }
    }

}