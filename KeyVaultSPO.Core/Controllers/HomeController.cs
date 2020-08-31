//===============================================================================
// Microsoft FastTrack for Azure
// Azure Key Vault Samples for SharePoint Online
//===============================================================================
// Copyright © Microsoft Corporation.  All rights reserved.
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
// OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
// LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
// FITNESS FOR A PARTICULAR PURPOSE.
//===============================================================================
using KeyVaultSPO.Core.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.KeyVault;
using Microsoft.Azure.KeyVault.Models;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace KeyVaultSPO.Core.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private IConfiguration _configuration;
        private AzureServiceTokenProvider _azureServiceTokenProvider;
        private KeyVaultClient _keyVaultClient;
        private AuthenticationManager _authenticationManager;
        private readonly string _listName = "ProjectList";

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;

            // Authenticate to Key Vault using the application's Managed Identity
            _azureServiceTokenProvider = new AzureServiceTokenProvider();
            _keyVaultClient = new KeyVaultClient(
                new KeyVaultClient.AuthenticationCallback(
                    _azureServiceTokenProvider.KeyVaultTokenCallback));

            _authenticationManager = new AuthenticationManager();
        }

        public async Task<IActionResult> Index()
        {
            List<Post> posted = new List<Post>();

            // Retrieve the certificate for the application credentials from Key Vault
            SecretBundle certificateSecret = await _keyVaultClient.GetSecretAsync(Environment.GetEnvironmentVariable("KEYVAULT_ENDPOINT"), "nickoftime-certificate");
            byte[] privateKeyBytes = Convert.FromBase64String(certificateSecret.Value);
            X509Certificate2 certificate = new X509Certificate2(privateKeyBytes, (string)null);

            // Authenticate to SPO using App only credentials and retrieve list data
            using (ClientContext clientContext = new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(Environment.GetEnvironmentVariable("SITE_URL"), Environment.GetEnvironmentVariable("CLIENT_ID"), Environment.GetEnvironmentVariable("TENANT"), certificate))
            {
                List projectList = clientContext.Web.Lists.GetByTitle(_listName);
                CamlQuery projectListQuery = new CamlQuery();
                projectListQuery.ViewXml = "<View><Query><OrderBy><FieldRef Name='Created' Ascending='FALSE'/></OrderBy></Query></View>";
                ListItemCollection projectListItems = projectList.GetItems(projectListQuery);
                clientContext.Load(projectListItems);
                clientContext.ExecuteQuery();
                foreach (ListItem p in projectListItems)
                {
                    Post post = MapListItemToPost(p);
                    posted.Add(post);
                }
            };

            return View("List", posted);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }


        private Post MapListItemToPost(ListItem p)
        {
            Post post = new Post();

            post.ID = p.Id;
            post.Title = p["Title"].ToString();
            post.Description = p["Description"].ToString();
            post.Type = p["Type"].ToString();
            if (p["EffortHours"] != null) post.EffortHours = Convert.ToInt32(p["EffortHours"]);
            if (p["EffortMinutes"] != null) post.EffortMinutes = Convert.ToInt32(p["EffortMinutes"]);
            if (p["StartDate"] != null) post.StartDate = Convert.ToDateTime(p["StartDate"]);
            if (p["EndDate"] != null) post.EndDate = Convert.ToDateTime(p["EndDate"]);
            post.ExpirationDate = Convert.ToDateTime(p["ExpirationDate"]);
            post.Location = p["Location"].ToString();
            FieldUserValue userField = (FieldUserValue)p["PostedBy"];
            post.PostedBy = userField.LookupValue;
            post.PostedByID = userField.LookupId;
            post.PostedByEmailAddress = userField.Email;
            post.Status = p["Status"].ToString();
            post.Skills = new List<string>();
            for (int i = 1; i < 11; i++)
            {
                if (p[string.Format("Skill{0}", i)] != null)
                {
                    post.Skills.Add(p[string.Format("Skill{0}", i)].ToString());
                }
            }

            return post;
        }
    }
}
