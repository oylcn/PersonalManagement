using MediatR;
using System;
using System.Collections.Generic;
using System.Text;

namespace PersonalManagement.Application.Features.Employees.Commands.DeleteEmployee
{
    public class DeleteEmployeeCommand : IRequest
    {
        public int EmployeeId { get; set; }
    }
}
