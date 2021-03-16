using PersonalManagement.BlazorApp.Services;
using PersonalManagement.BlazorApp.Services.Base;
using PersonalManagement.BlazorApp.ViewModels;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace PersonalManagement.BlazorApp.Contracts
{
    public interface IEmployeeDataService
    {
        Task<List<EmployeeListViewModel>> GetAllEmployees();
        Task<EmployeeDetailViewModel> GetEmployeeById(int id);
        Task<ApiResponse<int>> CreateEmployee(EmployeeDetailViewModel employeeDetailViewModel);
    }
}
