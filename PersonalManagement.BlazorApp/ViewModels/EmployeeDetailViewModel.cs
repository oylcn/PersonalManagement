using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace PersonalManagement.BlazorApp.ViewModels
{
    public class EmployeeDetailViewModel
    {
        public int EmployeeId { get; set; }

        [Required]
        public string FirstName { get; set; }

        [Required]
        public string LastName { get; set; }

        [Required]
        public string EmailAddress { get; set; }

        public string Mobile { get; set; }
      
        [Required]
        public int DepartmentId { get; set; }

        public DepartmentViewModel Department { get; set; }

        [Required]
        public int TitleId { get; set; }

        public TitleViewModel Title { get; set; }
    }
}
