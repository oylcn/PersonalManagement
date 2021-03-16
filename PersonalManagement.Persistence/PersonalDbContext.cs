using Microsoft.EntityFrameworkCore;
using PersonalManagement.Application.Contracts;
using PersonalManagement.Domain.Common;
using PersonalManagement.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PersonalManagement.Persistence
{
    public class PersonalDbContext : DbContext
    {
        private readonly ILoggedInUserService _loggedInUserService;
        public PersonalDbContext(DbContextOptions<PersonalDbContext> options) : base(options)
        {
        }

        public PersonalDbContext(DbContextOptions<PersonalDbContext> options, ILoggedInUserService loggedInUserService) : base(options)
        {
            _loggedInUserService = loggedInUserService;
        }

        public DbSet<Employee> Employees { get; set; }
        public DbSet<Department> Departments { get; set; }
        public DbSet<Title> Titles { get; set; }

        public override Task<int> SaveChangesAsync(CancellationToken cancellationToken = new CancellationToken())
        {
            foreach (var entry in ChangeTracker.Entries<AuditableEntity>())
            {
                switch (entry.State)
                {
                    case EntityState.Added:
                        entry.Entity.CreatedDate = DateTime.Now;
                        entry.Entity.CreatedBy = _loggedInUserService.UserId;
                        break;
                    case EntityState.Modified:
                        entry.Entity.LastModifiedDate = DateTime.Now;
                        entry.Entity.LastModifiedBy = _loggedInUserService.UserId;
                        break;
                }
            }
            return base.SaveChangesAsync(cancellationToken);
        }
    }
}
