//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Diagnostics;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
// TODO: Add directives for Graph and ADAL

namespace MailApp
{
    public class AuthenticationHelper
    {

        // These resources are added to the app.xaml file when the client application is registered.
        // The Client ID is used by the application to uniquely identify itself to the v2.0 authentication endpoint.
        static string clientId = App.Current.Resources["ida:ClientId"].ToString();
        static string aadInstance = App.Current.Resources["ida:AADInstance"].ToString();
        static string domain = App.Current.Resources["ida:Domain"].ToString();

        // Form the redirect URI that was registered in AAD by the Office Connected Services wizard.
        static string uriString = string.Format("ms-appx-web://microsoft.aad.brokerplugin/{0}", WebAuthenticationBroker.GetCurrentApplicationCallbackUri().Host).ToUpper();

        
        public static string TokenForUser = null;
        public static DateTimeOffset Expiration;

        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static GraphServiceClient GetAuthenticatedClient()
        {

            

            if (graphClient == null)
            {
                // Create Microsoft Graph client.
                try
                {
                    // TODO: Create the GraphServiceClient with the authentication provider.


                }

                catch (Exception ex)
                {
                    Debug.WriteLine("Could not create a graph client: " + ex.Message);
                }
            }

            return graphClient;
        }


        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> GetTokenForUserAsync()
        {
            // TODO: Create the authentication context based on the registered application's domain.
            

            // TODO: Recreate the redirectUri that was registered with Azure AD.
            

            // TODO: Authenticate the client. This leads to a prompt for credentials
            // and giving consent to the scopes specified in the Connected Services Wizard.
            // We specify here:
            // 1) Resource to access - in this case Microsoft Graph.
            // 2) Identify the client - this is how the scopes (capability) for consent are determined. 
            // 3) redirectUri - used for side loading during development
            // 4) PlatformParameters - only prompt for credentials when necessary, and we're not using SSO.
            

            // TODO: return the access token.
            
        }

        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static void SignOut()
        {
            //foreach (var user in IdentityClientApp.Users)
            //{
            //    user.SignOut();
            //}
            //graphClient = null;
            //TokenForUser = null;

        }
    }
}
