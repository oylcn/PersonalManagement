using AutoMapper;
using MediatR;
using PersonalManagement.Application.Contracts.Persistence;
using PersonalManagement.Application.Exceptions;
using PersonalManagement.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PersonalManagement.Application.Features.Employees.Commands.DeleteEmployee
{
    public class DeleteEmployeeCommandHandler : IRequestHandler<DeleteEmployeeCommand>
    {
        private readonly IAsyncRepository<Employee> _employeeRepository;
        private readonly IMapper _mapper;

        public DeleteEmployeeCommandHandler(IMapper mapper, IAsyncRepository<Employee> employeeRepository)
        {
            _mapper = mapper;
            _employeeRepository = employeeRepository;
        }

        public async Task<Unit> Handle(DeleteEmployeeCommand request, CancellationToken cancellationToken)
        {
            var employeeToDelete = await _employeeRepository.GetByIdAsync(request.EmployeeId);

            if (employeeToDelete == null)
            {
                throw new NotFoundException(nameof(Employee), request.EmployeeId);
            }

            await _employeeRepository.DeleteAsync(employeeToDelete);

            return Unit.Value;
        }
    }
}
