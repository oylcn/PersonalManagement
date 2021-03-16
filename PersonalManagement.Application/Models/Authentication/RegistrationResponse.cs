using System;
using System.Collections.Generic;
using System.Text;

namespace PersonalManagement.Application.Models.Authentication
{
    public class RegistrationResponse
    {
        public RegistrationResponse()
        {
            Success = true;
        }
        public string UserId { get; set; }
        public bool Success { get; set; }
        public List<string> ValidationErrors { get; set; }
    }
}
