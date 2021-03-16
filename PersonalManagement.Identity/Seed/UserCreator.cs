using Microsoft.AspNetCore.Identity;
using PersonalManagement.Identity.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace PersonalManagement.Identity.Seed
{
    public static class UserCreator
    {
        public static async Task<string> EnsureUser(UserManager<ApplicationUser> userManager)
        {
            var applicationUser = new ApplicationUser
            {
                FirstName = "Admin",
                LastName = "Admin",
                UserName = "Admin",
                Email = "admin@test.com",
                EmailConfirmed = true
            };

            var user = await userManager.FindByEmailAsync(applicationUser.Email);
            if (user == null)
            {
                await userManager.CreateAsync(applicationUser, "Passw0rd$");
            }
            return applicationUser.UserName;
        }

        public static async Task<IdentityResult> EnsureRole(RoleManager<IdentityRole> roleManager,UserManager<ApplicationUser> userManager,
                                                              string uid, string role)
        {
            if (!await roleManager.RoleExistsAsync(role))
            {
                await roleManager.CreateAsync(new IdentityRole(role));
            }

            var user = await userManager.FindByIdAsync(uid);

            if (user == null)
            {
                throw new Exception("The testUserPw password was probably not strong enough!");
            }

            IdentityResult IR = await userManager.AddToRoleAsync(user, role);
            return IR;
        }
    }
}
