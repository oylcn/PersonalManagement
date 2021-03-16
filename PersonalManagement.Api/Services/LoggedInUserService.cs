using Microsoft.AspNetCore.Http;
using PersonalManagement.Application.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;

namespace PersonalManagement.Api.Services
{
    public class LoggedInUserService : ILoggedInUserService
    {
        public LoggedInUserService(IHttpContextAccessor httpContextAccessor)
        {
            UserId = httpContextAccessor.HttpContext?.User?.FindFirstValue(ClaimTypes.NameIdentifier);
        }
        public string UserId { get; }
    }
}
