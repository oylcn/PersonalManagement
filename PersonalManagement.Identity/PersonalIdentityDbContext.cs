using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using PersonalManagement.Identity.Models;

namespace PersonalManagement.Identity
{
    public class PersonalIdentityDbContext : IdentityDbContext<ApplicationUser>
    {
        public PersonalIdentityDbContext(DbContextOptions<PersonalIdentityDbContext> options) : base(options)
        {
        }
    }

}
