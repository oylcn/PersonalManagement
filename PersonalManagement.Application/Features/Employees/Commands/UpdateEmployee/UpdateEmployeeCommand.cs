using MediatR;
using System;
using System.Collections.Generic;
using System.Text;

namespace PersonalManagement.Application.Features.Employees.Commands.UpdateEmployee
{
    public class UpdateEmployeeCommand : IRequest
    {
        public int EmployeeId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string EmailAddress { get; set; }
        public string Mobile { get; set; }
        public int TitleId { get; set; }
        public int DepartmentId { get; set; }
    }
}
