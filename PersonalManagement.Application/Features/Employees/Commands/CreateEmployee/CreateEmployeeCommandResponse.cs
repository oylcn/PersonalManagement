using PersonalManagement.Application.Responses;
using System;
using System.Collections.Generic;
using System.Text;

namespace PersonalManagement.Application.Features.Employees.Commands.CreateEmployee
{
    public class CreateEmployeeCommandResponse : BaseResponse
    {
        public CreateEmployeeCommandResponse() : base()
        {
        }

        public CreateEmployeeDto Employee { get; set; }
    }
}
