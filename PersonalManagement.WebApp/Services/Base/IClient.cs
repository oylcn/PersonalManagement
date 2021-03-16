using System.Net.Http;

namespace PersonalManagement.WebApp.Services
{
    public partial interface IClient
    {
        public HttpClient HttpClient { get; }

    }
}
