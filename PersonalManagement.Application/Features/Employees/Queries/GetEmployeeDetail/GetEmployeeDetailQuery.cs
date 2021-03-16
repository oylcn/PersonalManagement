using MediatR;

namespace PersonalManagement.Application.Features.Employees.Queries.GetEmployeeDetail
{
    public class GetEmployeeDetailQuery : IRequest<EmployeeDetailVm>
    {
        public int EmployeeId { get; set; }
    }
}
