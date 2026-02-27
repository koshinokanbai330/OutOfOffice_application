using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;

namespace OutOfOfficeAddin.Services
{
    /// <summary>
    /// Provides MSAL-based interactive token acquisition for Microsoft Graph.
    /// Configuration is read from appsettings.json that lives next to the add-in DLL.
    /// </summary>
    public class GraphAuthService
    {
        private const string GraphScope = "https://graph.microsoft.com/MailboxSettings.ReadWrite";

        private readonly string _clientId;
        private readonly string _tenantId;
        private readonly string _redirectUri;
        private IPublicClientApplication _app;

        public GraphAuthService()
        {
            var config = LoadConfig();
            _clientId = config["AzureAd"]["ClientId"]?.ToString() ?? string.Empty;
            _tenantId = config["AzureAd"]["TenantId"]?.ToString() ?? string.Empty;
            _redirectUri = config["AzureAd"]["RedirectUri"]?.ToString() ?? "http://localhost";
        }

        /// <summary>
        /// Acquires an access token interactively (or silently if a cached token exists).
        /// </summary>
        public async Task<string> AcquireTokenAsync()
        {
            var assemblyDir = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location) ?? string.Empty;
            var configPath = System.IO.Path.Combine(assemblyDir, "appsettings.json");

            if (string.IsNullOrWhiteSpace(_clientId) || string.IsNullOrWhiteSpace(_tenantId))
                throw new InvalidOperationException(
                    $"AzureAd:ClientId and AzureAd:TenantId must be set in '{configPath}' " +
                    "before using the OOF feature.");

            if (_app == null)
            {
                _app = PublicClientApplicationBuilder
                    .Create(_clientId)
                    .WithAuthority(AzureCloudInstance.AzurePublic, _tenantId)
                    .WithRedirectUri(_redirectUri)
                    .Build();
            }

            var scopes = new[] { GraphScope };

            // Try silent first (use cached token)
            try
            {
                var accounts = await _app.GetAccountsAsync();
                foreach (var account in accounts)
                {
                    try
                    {
                        var silentResult = await _app.AcquireTokenSilent(scopes, account).ExecuteAsync();
                        return silentResult.AccessToken;
                    }
                    catch (MsalUiRequiredException) { /* fall through to interactive */ }
                }
            }
            catch { /* fall through to interactive */ }

            // Interactive login
            var result = await _app
                .AcquireTokenInteractive(scopes)
                .WithPrompt(Prompt.SelectAccount)
                .ExecuteAsync();

            return result.AccessToken;
        }

        private static JObject LoadConfig()
        {
            var assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;
            var configPath = Path.Combine(assemblyDir, "appsettings.json");

            if (!File.Exists(configPath))
                return new JObject(new JProperty("AzureAd", new JObject()));

            var json = File.ReadAllText(configPath);
            return JObject.Parse(json);
        }
    }
}
