using AutoMapper;
using MediatR;
using PersonalManagement.Application.Contracts.Persistence;
using PersonalManagement.Application.Exceptions;
using PersonalManagement.Domain.Entities;
using System.Threading;
using System.Threading.Tasks;

namespace PersonalManagement.Application.Features.Employees.Queries.GetEmployeeDetail
{
    public class GetEmployeeDetailQueryHandler : IRequestHandler<GetEmployeeDetailQuery, EmployeeDetailVm>
    {
        private readonly IAsyncRepository<Employee> _employeeRepository;
        private readonly IAsyncRepository<Department> _departmentRepository;
        private readonly IAsyncRepository<Title> _titleRepository;
        private readonly IMapper _mapper;

        public GetEmployeeDetailQueryHandler(IMapper mapper, IAsyncRepository<Employee> employeeRepository, IAsyncRepository<Department> departmentRepository, IAsyncRepository<Title> titleRepository)
        {
            _mapper = mapper;
            _employeeRepository = employeeRepository;
            _departmentRepository = departmentRepository;
            _titleRepository = titleRepository;
        }
        public async Task<EmployeeDetailVm> Handle(GetEmployeeDetailQuery request, CancellationToken cancellationToken)
        {
            var @employee = await _employeeRepository.GetByIdAsync(request.EmployeeId);
            var employeeDetailDto = _mapper.Map<EmployeeDetailVm>(@employee);

            var department = await _departmentRepository.GetByIdAsync(@employee.DepartmentId);

            if (department == null)
            {
                throw new NotFoundException(nameof(employee), request.EmployeeId);
            }
            employeeDetailDto.Department = _mapper.Map<DepartmentDto>(department);

            var title = await _titleRepository.GetByIdAsync(@employee.TitleId);

            if (title == null)
            {
                throw new NotFoundException(nameof(employee), request.EmployeeId);
            }
            employeeDetailDto.Title = _mapper.Map<TitleDto>(title);

            return employeeDetailDto;
        }
    }
}
