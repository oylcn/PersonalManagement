using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using PersonalManagement.WebApp.Contracts;
using PersonalManagement.WebApp.Models;

namespace PersonalManagement.WebApp.Controllers
{
    public class AccountController : Controller
    {
        private readonly IAuthenticationService _authenticationService;

        public AccountController(IAuthenticationService authenticationService)
        {
            _authenticationService = authenticationService;
        }

        public IActionResult Index()
        {
            return RedirectToAction("Login", "Account");
        }
        public IActionResult Login()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(LoginViewModel loginViewModel) 
        {
            if (await _authenticationService.Authenticate(loginViewModel.Email, loginViewModel.Password))
            {
                return RedirectToAction("Index", "Home");
            }
            return RedirectToAction("Login", "Account", new { r = 4, m = "Kullanıcı adı veya şifreniz hatalıdır." });
        }
    }
}
