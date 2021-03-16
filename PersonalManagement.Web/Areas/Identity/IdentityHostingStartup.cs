using System;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PersonalManagement.Web.Data;

[assembly: HostingStartup(typeof(PersonalManagement.Web.Areas.Identity.IdentityHostingStartup))]
namespace PersonalManagement.Web.Areas.Identity
{
    public class IdentityHostingStartup : IHostingStartup
    {
        public void Configure(IWebHostBuilder builder)
        {
            builder.ConfigureServices((context, services) => {
                services.AddDbContext<PersonalManagementWebContext>(options =>
                    options.UseSqlServer(
                        context.Configuration.GetConnectionString("PersonalManagementWebContextConnection")));

                services.AddDefaultIdentity<ApplicationUser>(options => options.SignIn.RequireConfirmedAccount = true)
                    .AddEntityFrameworkStores<PersonalManagementWebContext>();
            });
        }
    }
}