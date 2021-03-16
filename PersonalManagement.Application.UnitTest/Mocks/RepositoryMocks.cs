using Moq;
using PersonalManagement.Application.Contracts.Persistence;
using PersonalManagement.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace PersonalManagement.Application.UnitTest.Mocks
{
    public class RepositoryMocks
    {
        public static Mock<IEmployeeRepository> GetEmployeeRepository()
        {          
            var employees = new List<Employee>
            {
                new Employee
                {
                    EmployeeId = 1,
                    FirstName = "Hasan",
                    LastName = "Serin",
                    EmailAddress = "hasan@mail.com",
                    TitleId = 1,
                    DepartmentId = 1,
                    Mobile = "",
                    CreatedBy = "admin",
                    CreatedDate = DateTime.Now,
                    LastModifiedBy = "admin",
                    LastModifiedDate = DateTime.Now                   
                },
                  new Employee
                {
                    EmployeeId = 2,
                    FirstName = "Veli",
                    LastName = "Serin",
                    EmailAddress = "veli@mail.com",
                    TitleId = 1,
                    DepartmentId = 1,
                    Mobile = "",
                    CreatedBy = "admin",
                    CreatedDate = DateTime.Now,
                    LastModifiedBy = "admin",
                    LastModifiedDate = DateTime.Now
                },
            };

            var mockEmployeeRepository = new Mock<IEmployeeRepository>();
            mockEmployeeRepository.Setup(repo => repo.ListAllAsync()).ReturnsAsync(employees);

            mockEmployeeRepository.Setup(repo => repo.AddAsync(It.IsAny<Employee>())).ReturnsAsync(
                (Employee employee) =>
                {
                    employees.Add(employee);
                    return employee;
                });

            return mockEmployeeRepository;
        }
    }
}
