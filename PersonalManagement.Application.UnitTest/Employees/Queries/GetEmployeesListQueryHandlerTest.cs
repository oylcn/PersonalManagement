using AutoMapper;
using Moq;
using PersonalManagement.Application.Contracts.Persistence;
using PersonalManagement.Application.Features.Employees.Queries.GetEmployeesList;
using PersonalManagement.Application.Profiles;
using PersonalManagement.Application.UnitTest.Mocks;
using PersonalManagement.Domain.Entities;
using Shouldly;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace PersonalManagement.Application.UnitTest.Employees.Queries
{
    public class GetEmployeesListQueryHandlerTest
    {
        private readonly IMapper _mapper;
        private readonly Mock<IEmployeeRepository> _mockEmployeeRepository;

        public GetEmployeesListQueryHandlerTest()
        {
            _mockEmployeeRepository = RepositoryMocks.GetEmployeeRepository();
            var configurationProvider = new MapperConfiguration(cfg =>
            {
                cfg.AddProfile<MappingProfile>();
            });

            _mapper = configurationProvider.CreateMapper();
        }

        [Fact]
        public async Task GetEmployeesListTest()
        {
            var handler = new GetEmployeesListQueryHandler(_mapper, _mockEmployeeRepository.Object);

            var result = await handler.Handle(new GetEmployeesListQuery(), CancellationToken.None);

            result.ShouldBeOfType<List<EmployeeListVm>>();

            result.Count.ShouldBe(4);
        }
    }
}
