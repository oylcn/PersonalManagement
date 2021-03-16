using Microsoft.EntityFrameworkCore;
using PersonalManagement.Application.Contracts.Persistence;
using PersonalManagement.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PersonalManagement.Persistence.Repositories
{
    public class EmployeeRepository : BaseRepository<Employee>, IEmployeeRepository
    {
        public EmployeeRepository(PersonalDbContext dbContext) : base(dbContext)
        {
        }

        public Task<bool> IsEmailAddressUnique(string emailAddress)
        {
            var matches = _dbContext.Employees.Any(e => e.EmailAddress.Equals(emailAddress));
            return Task.FromResult(matches);
        }

        public async Task<IReadOnlyList<Employee>> ListAllWithRelatedDataAsync()
        {
            return await _dbContext.Set<Employee>().Include("Department").Include("Title").ToListAsync();
        }
    }
}
