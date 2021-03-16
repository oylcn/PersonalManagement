using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PersonalManagement.WebApp.Contracts
{
    public interface IAuthenticationService
    {
        Task<bool> Authenticate(string email, string password);
        Task<bool> Register(string firstName, string lastName, string userName, string email, string password);
        Task Logout();
    }
}
