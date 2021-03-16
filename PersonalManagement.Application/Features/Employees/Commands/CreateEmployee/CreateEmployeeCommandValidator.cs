using FluentValidation;
using PersonalManagement.Application.Contracts.Persistence;
using PersonalManagement.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PersonalManagement.Application.Features.Employees.Commands.CreateEmployee
{
    public class CreateEmployeeCommandValidator : AbstractValidator<CreateEmployeeCommand>
    {
        private readonly IEmployeeRepository _employeeRepository;
        public CreateEmployeeCommandValidator(IEmployeeRepository employeeRepository)
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
                .EmailAddress(FluentValidation.Validators.EmailValidationMode.AspNetCoreCompatible).WithMessage("Email address is wrong.");

            RuleFor(p => p)
                .MustAsync(EmailAddressUnique)
                .WithMessage("The email address already exits.");

        }

        private async Task<bool> EmailAddressUnique(CreateEmployeeCommand e, CancellationToken token)
        {
            return !(await _employeeRepository.IsEmailAddressUnique(e.EmailAddress));
        }

    }
}
