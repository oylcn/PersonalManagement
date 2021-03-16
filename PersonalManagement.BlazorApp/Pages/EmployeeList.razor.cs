using PersonalManagement.BlazorApp.Contracts;
using PersonalManagement.BlazorApp.ViewModels;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;

namespace PersonalManagement.BlazorApp.Pages
{
    public partial class EmployeeList
    {
        [Inject]
        public IEmployeeDataService EmployeeDataService { get; set; }

        [Inject]
        public NavigationManager NavigationManager { get; set; }

        public ICollection<EmployeeListViewModel> Employees { get; set; }

        [Inject]
        public IJSRuntime JSRuntime { get; set; }

        protected async override Task OnInitializedAsync()
        {
            Employees = await EmployeeDataService.GetAllEmployees();
        }

        protected void AddNewEmployee()
        {
            NavigationManager.NavigateTo("/employeedetails");
        }

        [Inject]
        public HttpClient HttpClient { get; set; }
    }
}
