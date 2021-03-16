using AutoMapper;
using Blazored.LocalStorage;
using PersonalManagement.BlazorApp.Contracts;
using PersonalManagement.BlazorApp.Services.Base;
using PersonalManagement.BlazorApp.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PersonalManagement.BlazorApp.Services
{
    public class EmployeeDataService: BaseDataService, IEmployeeDataService
    {
        
        private readonly IMapper _mapper;

        public EmployeeDataService(IClient client, IMapper mapper, ILocalStorageService localStorage) : base(client, localStorage)
        {
            _mapper = mapper;
        }

        public async Task<List<EmployeeListViewModel>> GetAllEmployees()
        {
            var allEmployees = await _client.GetEmployeesAsync();
            var mappedEmployees = _mapper.Map<ICollection<EmployeeListViewModel>>(allEmployees);
            return mappedEmployees.ToList();
        }

        public async Task<EmployeeDetailViewModel> GetEmployeeById(int id)
        {
            var selectedEmployee = await _client.GetEmployeeByIdAsync(id);
            var mappedEmployee = _mapper.Map<EmployeeDetailViewModel>(selectedEmployee);
            return mappedEmployee;
        }

        public async Task<ApiResponse<int>> CreateEmployee(EmployeeDetailViewModel employeeDetailViewModel)
        {
            try
            {
                CreateEmployeeCommand createEmployeeCommand = _mapper.Map<CreateEmployeeCommand>(employeeDetailViewModel);
                var response = await _client.AddEmployeeAsync(createEmployeeCommand);
                return new ApiResponse<int>() { Data = response.Employee.EmployeeId, Success = true };
            }
            catch (ApiException ex)
            {
                return ConvertApiExceptions<int>(ex);
            }
        }

        public async Task<ApiResponse<Guid>> UpdateEmployee(EmployeeDetailViewModel employeeDetailViewModel)
        {
            try
            {
                UpdateEmployeeCommand updateEmployeeCommand = _mapper.Map<UpdateEmployeeCommand>(employeeDetailViewModel);
                await _client.UpdateEmployeeAsync(updateEmployeeCommand);
                return new ApiResponse<Guid>() { Success = true };
            }
            catch (ApiException ex)
            {
                return ConvertApiExceptions<Guid>(ex);
            }
        }

        public async Task<ApiResponse<int>> DeleteEmployee(int id)
        {
            try
            {
                await _client.DeleteEmployeeAsync(id);
                return new ApiResponse<int>() { Success = true };
            }
            catch (ApiException ex)
            {
                return ConvertApiExceptions<int>(ex);
            }
        }
    }
}
