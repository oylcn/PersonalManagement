using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using PersonalManagement.Application.Contracts.Identity;
using PersonalManagement.Application.Models.Authentication;

namespace PersonalManagement.Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AccountController : ControllerBase
    {
        private readonly IAuthenticationService _authenticationService;
        public AccountController(IAuthenticationService authenticationService)
        {
            _authenticationService = authenticationService;
        }

        [HttpPost("authenticate")]
        public async Task<ActionResult<AuthenticationResponse>> AuthenticateAsync(AuthenticationRequest request)
        {
            return Ok(await _authenticationService.AuthenticateAsync(request));
        }

        [HttpPost("register")]
        public async Task<ActionResult<RegistrationResponse>> RegisterAsync(RegistrationRequest request)
        {
            return Ok(await _authenticationService.RegisterAsync(request));
        }
    }
}
