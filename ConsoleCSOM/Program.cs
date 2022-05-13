﻿using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
namespace ConsoleCSOM
{
    class SharepointInfo
    {
        public string SiteUrl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper())
                {
                    ClientContext ctx = GetContext(clientContextHelper);
                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();
                    Console.WriteLine($"Site {ctx.Web.Title}");
                    await UpdatePropertyForUser(ctx, "Precio-Nickname", new List<string> { "tu.nguyen.dev@devtusturu.onmicrosoft.com", "tu.nguyen.tester@devtusturu.onmicrosoft.com", "tu.nguyen.anh@devtusturu.onmicrosoft.com" });
                    //await LoadUser(ctx);

                }
                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }

        private static async Task LoadUser(ClientContext ctx)
        {
            UserCollection users = ctx.Web.SiteUsers;
            // Load the user collection  
            ctx.Load(users);
            // Execute the query  
            await ctx.ExecuteQueryAsync();
            // Check if the owners are not null  
            if (users != null)
            {
                // Loop through all the users  
                foreach (var user in users)
                {
                    // Check if the users email is not empty  
                    // O365 group added to owners groups will be displayed for Modern sites  
                    // If you want to retrieve only the users from default user group then check if principal type is user  
                    if (user.PrincipalType.ToString() == "User" && !String.IsNullOrEmpty(user.Email))
                    {
                        PeopleManager peopleManager = new PeopleManager(ctx);
                        PersonProperties personProperties = peopleManager.GetPropertiesFor(user.LoginName);
                        ctx.Load(personProperties);
                        await ctx.ExecuteQueryAsync();
                        Console.WriteLine(
                           $"Account Name: {personProperties.UserProfileProperties["AccountName"]}\n" +
                           $"Email: {personProperties.UserProfileProperties["WorkEmail"]}\n" +
                           $"Nickname: {personProperties.UserProfileProperties["Precio-Nickname"]}\n" +
                           $"Cities: {personProperties.UserProfileProperties["Precio-Cities"]}\n" +
                           $"Job Title: {personProperties.UserProfileProperties["Precio-JobTitle"]}"
                           );

                    }
                }
            }
        }
        private static async Task UpdatePropertyForUser(ClientContext ctx, string propertyName, List<string> userEmailList)
        {
            foreach (string userEmail in userEmailList)
            {
                User user = ctx.Web.EnsureUser(userEmail);
                ctx.Load(user);
                await ctx.ExecuteQueryAsync();
                if (user != null)
                {
                    try
                    {
                        PeopleManager peopleManager = new PeopleManager(ctx);
                        peopleManager.SetSingleValueProfileProperty(user.LoginName, propertyName, "Test");
                        ctx.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
        }
    }
}
