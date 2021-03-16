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

namespace PersonalManagement.Application.Features.Employees.Commands.UpdateEmployee
{
    public class UpdateEmployeeCommandHandler : IRequestHandler<UpdateEmployeeCommand>
    {
        private readonly IAsyncRepository<Employee> _employeeRepository;
        private readonly IMapper _mapper;

        public UpdateEmployeeCommandHandler(IMapper mapper, IAsyncRepository<Employee> employeeRepository)
        {
            _mapper = mapper;
            _employeeRepository = employeeRepository;
        }

        public async Task<Unit> Handle(UpdateEmployeeCommand request, CancellationToken cancellationToken)
        {
            var employeeToUpdate = await _employeeRepository.GetByIdAsync(request.EmployeeId);

            if(employeeToUpdate == null)
            {
                throw new NotFoundException(nameof(Employee), request.EmployeeId);
            }

            var validator = new UpdateEmployeeCommandValidator(_employeeRepository);
            var validationResult = await validator.ValidateAsync(request);

            if (validationResult.Errors.Count > 0 )
                throw new ValidationException(validationResult);

            _mapper.Map(request, employeeToUpdate, typeof(UpdateEmployeeCommand), typeof(Employee));

            await _employeeRepository.UpdateAsync(employeeToUpdate);

            return Unit.Value;
        }
    }
}
