//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Diagnostics;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
// TODO: Add directives for Graph and ADAL
using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace MailApp
{
    public class AuthenticationHelper
    {
        // These resources are added to the app.xaml file when the client application is registered.
        static string clientId = App.Current.Resources["ida:ClientId"].ToString();
        static string aadInstance = App.Current.Resources["ida:AADInstance"].ToString();
        static string domain = App.Current.Resources["ida:Domain"].ToString();

        // Form the redirect URI that was registered in AAD by the Office 365 Connected Services wizard.
        static string uriString = string.Format("ms-appx-web://microsoft.aad.brokerplugin/{0}", WebAuthenticationBroker.GetCurrentApplicationCallbackUri().Host).ToUpper();
        
        public static string TokenForUser = null;
        public static DateTimeOffset Expiration;

        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId.
        public static GraphServiceClient GetAuthenticatedClient()
        {
            if (graphClient == null)
            {
                // Create Microsoft Graph client.
                try
                {
                    // TODO: Create the GraphServiceClient with the authentication provider.
                    graphClient = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                        new DelegateAuthenticationProvider(
                            async (requestMessage) =>
                            {
                                var token = await GetTokenForUserAsync();
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                            }));

                    return graphClient;
                }

                catch (Exception ex)
                {
                    Debug.WriteLine("Could not create a graph client: " + ex.Message);
                }
            }

            return graphClient;
        }

        private static async Task<string> GetTokenForUserAsync()
        {
            // TODO: Create the authentication context based on the registered application's domain.
            var authContext = new AuthenticationContext(aadInstance + domain, true);

            // TODO: Recreate the redirectUri that was registered with Azure AD.
            Uri redirectUri = new Uri(uriString);

            // TODO: Authenticate the client. This leads to a prompt for credentials
            // and giving consent to the scopes specified in the Connected Services Wizard.
            // We specify here:
            // 1) Resource to access - in this case Microsoft Graph.
            // 2) Identify the client - this is how the scopes (capability) for consent are determined. 
            // 3) redirectUri - used for side loading during development
            // 4) PlatformParameters - only prompt for credentials when necessary, and we're not using SSO.
            var authResult = await authContext.AcquireTokenAsync("https://graph.microsoft.com",
                                                                 clientId,
                                                                 redirectUri,
                                                                 new PlatformParameters(PromptBehavior.Auto, false));

            // TODO: return the access token.
            return authResult.AccessToken;
        }
        

        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static void SignOut()
        {
            throw new NotImplementedException("Implement sign out");
        }
    }
}
