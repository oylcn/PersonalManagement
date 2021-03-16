using Microsoft.AspNetCore.Http;
using PersonalManagement.WebApp.Services;
using PersonalManagement.WebApp.Services.Base;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace PersonalManagement.WebApp.Services.Base
{
    public class BaseDataService
    {
        
        protected IClient _client;
        private readonly IHttpContextAccessor _httpContextAccessor;
        public BaseDataService(IClient client, IHttpContextAccessor httpContextAccessor)
        {
            _client = client;
            _httpContextAccessor = httpContextAccessor;
        }

        protected ApiResponse<Guid> ConvertApiExceptions<Guid>(ApiException ex)
        {
            if (ex.StatusCode == 400)
            {
                return new ApiResponse<Guid>() { Message = "Validation errors have occured.", ValidationErrors = ex.Response, Success = false };
            }
            else if (ex.StatusCode == 404)
            {
                return new ApiResponse<Guid>() { Message = "The requested item could not be found.", Success = false };
            }
            else
            {
                return new ApiResponse<Guid>() { Message = "Something went wrong, please try again.", Success = false };
            }
        }

        protected void AddBearerToken()
        {
            var sessionHelper = new SessionHelper(_httpContextAccessor);
            var token = sessionHelper.GetString("token");
            if (!string.IsNullOrEmpty(token))
                _client.HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            //if (await _localStorage.ContainKeyAsync("token"))
            //    _client.HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await _localStorage.GetItemAsync<string>("token"));
        }
    }
}
