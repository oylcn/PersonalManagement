using System;
using System.Collections.Generic;
using System.Text;

namespace PersonalManagement.Application.Features.Employees.Queries.GetEmployeeDetail
{
    public class EmployeeDetailVm
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string EmailAddress { get; set; }
        public string Mobile { get; set; }
        public int TitleId { get; set; }
        public TitleDto Title { get; set; }
        public int DepartmentId { get; set; }
        public DepartmentDto Department { get; set; }
    }
}
