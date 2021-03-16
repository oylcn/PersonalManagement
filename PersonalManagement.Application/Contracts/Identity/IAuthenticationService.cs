using PersonalManagement.Application.Models.Authentication;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace PersonalManagement.Application.Contracts.Identity
{
    public interface IAuthenticationService
    {
        Task<AuthenticationResponse> AuthenticateAsync(AuthenticationRequest request);
        Task<RegistrationResponse> RegisterAsync(RegistrationRequest request);
    }
}
