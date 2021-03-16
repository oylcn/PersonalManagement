using AutoMapper;
using Blazored.LocalStorage;
using PersonalManagement.BlazorApp.Auth;
using PersonalManagement.BlazorApp.Contracts;
using PersonalManagement.BlazorApp.Services;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;

namespace PersonalManagement.BlazorApp
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var builder = WebAssemblyHostBuilder.CreateDefault(args);
            builder.RootComponents.Add<App>("app");

            builder.Services.AddAutoMapper(Assembly.GetExecutingAssembly());
            builder.Services.AddBlazoredLocalStorage();

            builder.Services.AddAuthorizationCore();
            builder.Services.AddScoped<AuthenticationStateProvider, CustomAuthenticationStateProvider>();


            builder.Services.AddSingleton(new HttpClient
            {
                BaseAddress = new Uri("https://localhost:44331")
            });

            builder.Services.AddHttpClient<IClient, Client>(client => client.BaseAddress = new Uri("https://localhost:44331"));

            builder.Services.AddScoped<IEmployeeDataService, EmployeeDataService>();
            builder.Services.AddScoped<IAuthenticationService, AuthenticationService>();

            await builder.Build().RunAsync();
        }
    }
}
