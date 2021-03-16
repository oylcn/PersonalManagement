using System;
using System.Collections.Generic;
using System.Text;

namespace PersonalManagement.Application.Contracts
{
    public interface ILoggedInUserService
    {
        string UserId { get; }
    }
}
