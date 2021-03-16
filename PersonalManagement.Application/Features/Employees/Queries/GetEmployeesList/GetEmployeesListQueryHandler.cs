using AutoMapper;
using MediatR;
using PersonalManagement.Application.Contracts.Persistence;
using PersonalManagement.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PersonalManagement.Application.Features.Employees.Queries.GetEmployeesList
{
    public class GetEmployeesListQueryHandler : IRequestHandler<GetEmployeesListQuery, List<EmployeeListVm>>
    {
        private readonly IEmployeeRepository _employeeRepository;
        private readonly IMapper _mapper;

        public GetEmployeesListQueryHandler(IMapper mapper, IEmployeeRepository employeeRepository)
        {
            _mapper = mapper;
            _employeeRepository = employeeRepository;
        }
        public async Task<List<EmployeeListVm>> Handle(GetEmployeesListQuery request, CancellationToken cancellationToken)
        {
            var allEmployees = (await _employeeRepository.ListAllWithRelatedDataAsync()).OrderByDescending(a=>a.CreatedDate);
            return _mapper.Map<List<EmployeeListVm>>(allEmployees);
        }
    }
}
