using AutoMapper;
using MediatR;
using Microsoft.Extensions.Logging;
using PersonalManagement.Application.Contracts.Persistence;
using PersonalManagement.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PersonalManagement.Application.Features.Employees.Commands.CreateEmployee
{
    public class CreateEmployeeCommandHandler : IRequestHandler<CreateEmployeeCommand, CreateEmployeeCommandResponse>
    {
        private readonly IEmployeeRepository _employeeRepository;
        private readonly IMapper _mapper;
        public CreateEmployeeCommandHandler(IMapper mapper, IEmployeeRepository employeeRepository, ILogger<CreateEmployeeCommandHandler> logger)
        {
            _mapper = mapper;
            _employeeRepository = employeeRepository;
        }
        public async Task<CreateEmployeeCommandResponse> Handle(CreateEmployeeCommand request, CancellationToken cancellationToken)
        {
            var createEmployeeCommandResponse = new CreateEmployeeCommandResponse();

            var validator = new CreateEmployeeCommandValidator(_employeeRepository);
            var validationResult = await validator.ValidateAsync(request);

            if (validationResult.Errors.Count > 0)
            {
                createEmployeeCommandResponse.Success = false;
                createEmployeeCommandResponse.ValidationErrors = new List<string>();
                foreach (var error in validationResult.Errors)
                {
                    createEmployeeCommandResponse.ValidationErrors.Add(error.ErrorMessage);
                }
            }
            if (createEmployeeCommandResponse.Success)
            {
                var employee = _mapper.Map<Employee>(request);
                employee = await _employeeRepository.AddAsync(employee);
                createEmployeeCommandResponse.Employee = _mapper.Map<CreateEmployeeDto>(employee);
            }

            return createEmployeeCommandResponse;
        }
    }
}
