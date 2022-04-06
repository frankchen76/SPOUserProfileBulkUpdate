using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.SharePoint.Client;
using PnP.Core.Auth;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Services;
using PnP.Core.Services.Builder.Configuration;
using PnP.Framework;
using SPOUserProfileBulkUpdate;
using System.Net;
using System.Security.Cryptography.X509Certificates;

// See https://aka.ms/new-console-template for more information
// Creates and configures the host
var host = Host.CreateDefaultBuilder()
    .ConfigureServices((context, services) =>
    {
        // Add PnP Core SDK
        services.AddPnPCore(options =>
        {
            string certFile = @"[pfx-path]";
            string clientId = "[client-id]";
            string tenantId = "[tenant-id]";
            X509Certificate2 certficate = new X509Certificate2(certFile, "[password]");

            options.DefaultAuthenticationProvider = new PnP.Core.Auth.X509CertificateAuthenticationProvider(clientId, tenantId, certficate);
            // Configure the interactive authentication provider as default
            //options.DefaultAuthenticationProvider = new InteractiveAuthenticationProvider()
            //{
            //    ClientId = clientId,
            //    TenantId= tenantId,
            //    RedirectUri = new Uri("http://localhost")
            //};
        });
    })
    .UseConsoleLifetime()
    .Build();

// Start the host
await host.StartAsync();

string siteUrl = "https://m365x725618.sharepoint.com/sites/FrankCommunication1";

using (var scope = host.Services.CreateScope())
{
    // Ask an IPnPContextFactory from the host
    var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();
    
    // Create a PnPContext
    using (var context = await pnpContextFactory.CreateAsync(new Uri(siteUrl)))
    {
        
        // Load the Title property of the site's root web
        await context.Web.LoadAsync(p => p.Title);
        Console.WriteLine($"The title of the web is {context.Web.Title}");

        

        using (ClientContext csomContext = PnPCoreSdk.Instance.GetClientContext(context))
        {
            // Use CSOM to load the web title
            csomContext.Load(csomContext.Web, p => p.Title);
            csomContext.ExecuteQueryRetry();

            UPSSync sync = new UPSSync();
            string jsonFile = "https://m365x725618.sharepoint.com/sites/FrankCommunication1/Shared Documents/UserProfileValues.json";
            sync.QueueUserProfileJob(csomContext, jsonFile);
        }
    }
}

