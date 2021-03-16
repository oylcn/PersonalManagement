using System;
using System.Collections.Generic;
using System.Text;

namespace PersonalManagement.Application.Features.Employees.Queries.GetEmployeesList
{
    public class EmployeeListVm
    {
        public int EmployeeId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string EmailAddress { get; set; }
        public string Mobile { get; set; }
        public string TitleName { get; set; }
        public string DepartmentName { get; set; }
    }
}
