using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;

namespace PluginSimulator.Authentication
{
    public static class AuthManager
    {
        public static IOrganizationService Connect(string environmentUrl, string connectionString = null)
        {
            CrmServiceClient client;

            if (!string.IsNullOrEmpty(connectionString))
            {
                client = new CrmServiceClient(connectionString);
            }
            else
            {
                // Interactive browser-based authentication
                // Uses default Azure AD app registration for dev tools
                var connStr = $"AuthType=OAuth;" +
                              $"Url={environmentUrl};" +
                              $"AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;" +
                              $"RedirectUri=app://58145B91-0C36-4500-8554-080854F2AC97;" +
                              $"LoginPrompt=Auto";
                client = new CrmServiceClient(connStr);
            }

            if (!client.IsReady)
            {
                throw new Exception($"Failed to connect to {environmentUrl}: {client.LastCrmError}");
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"  Connected to: {client.ConnectedOrgFriendlyName}");
            Console.WriteLine($"  Organization: {client.ConnectedOrgUniqueName}");
            Console.WriteLine($"  User: {client.OAuthUserId}");
            Console.ResetColor();

            return client;
        }
    }
}
