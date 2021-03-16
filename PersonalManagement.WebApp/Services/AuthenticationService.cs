using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Http;
using PersonalManagement.WebApp.Auth;
using PersonalManagement.WebApp.Contracts;
using PersonalManagement.WebApp.Services.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace PersonalManagement.WebApp.Services
{
    public class AuthenticationService : BaseDataService, IAuthenticationService
    {
        private readonly AuthenticationStateProvider _authenticationStateProvider;
        private readonly IHttpContextAccessor _httpContextAccessor;
        public AuthenticationService(IClient client, IHttpContextAccessor httpContextAccessor, AuthenticationStateProvider authenticationStateProvider) : base(client, httpContextAccessor)
        {
            _authenticationStateProvider = authenticationStateProvider;
            _httpContextAccessor = httpContextAccessor;
        }

        public async Task<bool> Authenticate(string email, string password)
        {
            try
            {
                AuthenticationRequest authenticationRequest = new AuthenticationRequest() { Email = email, Password = password };
                var authenticationResponse = await _client.AuthenticateAsync(authenticationRequest);

                if (authenticationResponse.Token != string.Empty)
                {
                    var sessionHelper = new SessionHelper(_httpContextAccessor);
                    sessionHelper.SetString("token", authenticationResponse.Token);
                    ((CustomAuthenticationStateProvider)_authenticationStateProvider).SetUserAuthenticated(email);
                    _client.HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("bearer", authenticationResponse.Token);
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        public async Task<bool> Register(string firstName, string lastName, string userName, string email, string password)
        {
            RegistrationRequest registrationRequest = new RegistrationRequest() { FirstName = firstName, LastName = lastName, Email = email, UserName = userName, Password = password };
            var response = await _client.RegisterAsync(registrationRequest);

            if (!string.IsNullOrEmpty(response.UserId))
            {
                return true;
            }
            return false;
        }

        public async Task Logout()
        {
            var sessionHelper = new SessionHelper(_httpContextAccessor);
            await Task.Run(() => sessionHelper.RemoveItem("token"));
            ((CustomAuthenticationStateProvider)_authenticationStateProvider).SetUserLoggedOut();
            _client.HttpClient.DefaultRequestHeaders.Authorization = null;
        }
    }
}
