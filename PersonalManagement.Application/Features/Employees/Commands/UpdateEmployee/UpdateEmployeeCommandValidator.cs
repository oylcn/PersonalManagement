using FluentValidation;
using PersonalManagement.Application.Contracts.Persistence;
using PersonalManagement.Application.Features.Employees.Commands.CreateEmployee;
using PersonalManagement.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PersonalManagement.Application.Features.Employees.Commands.UpdateEmployee
{
    class UpdateEmployeeCommandValidator : AbstractValidator<UpdateEmployeeCommand>
    {
        private readonly IAsyncRepository<Employee> _employeeRepository;
        public UpdateEmployeeCommandValidator(IAsyncRepository<Employee> employeeRepository)
        {
            _employeeRepository = employeeRepository;

            RuleFor(p => p.FirstName)
                .NotEmpty().WithMessage("{FirstName} is required.")
                .NotNull()
                .MaximumLength(50).WithMessage("{PropertyName} must not exceed 50 characters.");

            RuleFor(p => p.LastName)
                .NotEmpty().WithMessage("{LastName} is required.")
                .NotNull()
                .MaximumLength(50).WithMessage("{PropertyName} must not exceed 50 characters.");

            RuleFor(p => p.EmailAddress)
                .NotEmpty().WithMessage("{EmailAddress} is required.")
                .NotNull()
                .EmailAddress(FluentValidation.Validators.EmailValidationMode.AspNetCoreCompatible)
                .WithMessage("Email address is wrong.");

        }
    }
}