using AutoMapper;
using PersonalManagement.BlazorApp.Services;
using PersonalManagement.BlazorApp.ViewModels;

namespace PersonalManagement.BlazorApp.Profiles
{
    public class Mappings : Profile
    {
        public Mappings()
        {
            //Vms are coming in from the API, ViewModel are the local entities in Blazor
            CreateMap<EmployeeListVm, EmployeeListViewModel>().ReverseMap();
            CreateMap<EmployeeDetailVm, EmployeeDetailViewModel>().ReverseMap();

            CreateMap<EmployeeDetailViewModel, CreateEmployeeCommand>().ReverseMap();
            CreateMap<EmployeeDetailViewModel, UpdateEmployeeCommand>().ReverseMap();


            CreateMap<DepartmentDto, DepartmentViewModel>().ReverseMap();
            CreateMap<TitleDto, TitleViewModel>().ReverseMap();
        }
    }
}
