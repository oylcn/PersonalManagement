using System.Net.Http;

namespace PersonalManagement.BlazorApp.Services
{
    public partial interface IClient
    {
        public HttpClient HttpClient { get; }

    }
}
