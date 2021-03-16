using PersonalManagement.Domain.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace PersonalManagement.Domain.Entities
{
    public class Employee : AuditableEntity
    {
        public int EmployeeId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string EmailAddress { get; set; }
        public string Mobile { get; set; }
        public int TitleId { get; set; }
        public int DepartmentId { get; set; }
        public Department Department { get; set; }
        public Title Title { get; set; }

    }
}
