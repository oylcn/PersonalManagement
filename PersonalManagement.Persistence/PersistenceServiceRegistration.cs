using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PersonalManagement.Application.Contracts.Persistence;
using PersonalManagement.Persistence.Repositories;
using System;
using System.Collections.Generic;
using System.Text;

namespace PersonalManagement.Persistence
{
    public static class PersistenceServiceRegistration
    {
        public static IServiceCollection AddPersistenceServices(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddDbContext<PersonalDbContext>(options =>
                options.UseSqlServer(configuration.GetConnectionString("PersonalManagementConnectionString")));

            services.AddScoped(typeof(IAsyncRepository<>), typeof(BaseRepository<>));
            services.AddScoped<IEmployeeRepository, EmployeeRepository>();

            return services;
        }
    }
}
