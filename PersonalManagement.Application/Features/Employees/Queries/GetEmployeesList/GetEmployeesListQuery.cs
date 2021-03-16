using MediatR;
using System;
using System.Collections.Generic;
using System.Text;

namespace PersonalManagement.Application.Features.Employees.Queries.GetEmployeesList
{
    public class GetEmployeesListQuery : IRequest<List<EmployeeListVm>>
    {
    }
}
