using AutoMapper;
using PersonalManagement.Application.Features.Employees.Commands.CreateEmployee;
using PersonalManagement.Application.Features.Employees.Commands.UpdateEmployee;
using PersonalManagement.Application.Features.Employees.Queries.GetEmployeesList;
using PersonalManagement.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace PersonalManagement.Application.Profiles
{
    public class MappingProfile : Profile
    {
        public MappingProfile()
        {
            CreateMap<Employee, CreateEmployeeCommand>().ReverseMap();
            CreateMap<Employee, UpdateEmployeeCommand>().ReverseMap();
            CreateMap<Employee, EmployeeListVm>()
                .ForMember(
                    dest => dest.DepartmentName,
                    opt => opt.MapFrom(a => a.Department.Description)
                    )
                .ForMember(
                    dest => dest.TitleName,
                    opt => opt.MapFrom(a => a.Title.Description)
                    )
                .ReverseMap();
        }
    }
}
