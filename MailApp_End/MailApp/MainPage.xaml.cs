//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.ApplicationModel.Resources;
using Windows.ApplicationModel.Resources.Core;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using Windows.Storage;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace MailApp
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        private string _mailAddress;
        private string _displayName = null;
        private MailHelper _mailHelper = new MailHelper();
        public static ApplicationDataContainer _settings = ApplicationData.Current.RoamingSettings;

        public MainPage()
        {
            this.InitializeComponent();
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            // Developer code - if you haven't registered the app yet, we warn you. 
            if (!App.Current.Resources.ContainsKey("ida:ClientId"))
            {
                InfoText.Text = "Register your application by using Connected Services.";
                ConnectButton.IsEnabled = false;
            }
            else
            {
                InfoText.Text = "Tap the following button to connect to your work or school account.";
                ConnectButton.IsEnabled = true;
            }
        }

        /// <summary>
        /// Signs in the current user when you select the Connect button.
        /// </summary>
        /// <returns></returns>
        public async Task<bool> SignInCurrentUserAsync()
        {
            // TODO: Get an authenticated GraphServiceClient.
            var graphClient = AuthenticationHelper.GetAuthenticatedClient();

            if (graphClient != null)
            {
                // TODO: Call Microsoft Graph to get information about the authenticated user.
                var user = await graphClient.Me.Request().GetAsync();
                string userId = user.Id;
                _mailAddress = user.UserPrincipalName;
                _displayName = user.DisplayName;
                return true;
            }
            else
            {
                return false;
            }

        }


        private async void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            ProgressBar.Visibility = Visibility.Visible;
            if (await SignInCurrentUserAsync())
            {
                InfoText.Text = "Hi " + _displayName + "," + Environment.NewLine + "You're now connected to Microsoft Graph. Tap send mail to send a message from your account using Microsoft Graph.";
                MailButton.IsEnabled = true;
                EmailAddressBox.IsEnabled = true;
                ConnectButton.Visibility = Visibility.Collapsed;
                DisconnectButton.Visibility = Visibility.Visible;
                EmailAddressBox.Text = _mailAddress;
            }
            else
            {
                InfoText.Text = "Oops, couldn't connect to O365. ";
            }

            ProgressBar.Visibility = Visibility.Collapsed;
        }

        private async void MailButton_Click(object sender, RoutedEventArgs e)
        {
            _mailAddress = EmailAddressBox.Text;
            ProgressBar.Visibility = Visibility.Visible;
            MailStatus.Text = string.Empty;
            try
            {
                await _mailHelper.ComposeAndSendMailAsync("Mail subject", ComposePersonalizedMail(_displayName), _mailAddress);
                MailStatus.Text = string.Format("Success", _mailAddress);
            }
            catch (Exception)
            {
                MailStatus.Text = "Error";
            }
            finally
            {
                ProgressBar.Visibility = Visibility.Collapsed;
            }

        }

        // <summary>
        // Personalizes the email.
        // </summary>
        public static string ComposePersonalizedMail(string userName)
        {
            return $"Hello {userName}, this is an email.";
        }

        private void Disconnect_Click(object sender, RoutedEventArgs e)
        {
            ProgressBar.Visibility = Visibility.Visible;
            AuthenticationHelper.SignOut();
            ProgressBar.Visibility = Visibility.Collapsed;
            MailButton.IsEnabled = false;
            EmailAddressBox.IsEnabled = false;
            ConnectButton.Visibility = Visibility.Visible;
            InfoText.Text = "Disconnect";
            this._displayName = null;
            this._mailAddress = null;
        }
    }
}
