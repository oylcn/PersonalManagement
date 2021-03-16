using PersonalManagement.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace PersonalManagement.Application.Contracts.Persistence
{
    public interface IEmployeeRepository : IAsyncRepository<Employee>
    {
        Task<bool> IsEmailAddressUnique(string emailAddress);
        Task<IReadOnlyList<Employee>> ListAllWithRelatedDataAsync();
    }
}
