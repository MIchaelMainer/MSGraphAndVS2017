using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace MailApp
{
    class MailHelper
    {
        /// <summary>
        /// Compose and send a new email.
        /// </summary>
        /// <param name="subject">The subject line of the email.</param>
        /// <param name="bodyContent">The body of the email.</param>
        /// <param name="recipients">A semicolon-separated list of email addresses.</param>
        /// <returns></returns>
        public async Task ComposeAndSendMailAsync(string subject,
                                                            string bodyContent,
                                                            string recipients)
        {
            // Add the recipient for the email.
            List<Recipient> recipientList = new List<Recipient>();
            recipientList.Add(new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipients.Trim()
                }
            });

            try
            {
                // TODO: Get an authenticated GraphServiceClient.
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();

                // Form the email that we'll send
                var email = new Message()
                {
                    Body = new ItemBody
                    {
                        Content = bodyContent,
                        ContentType = BodyType.Html
                    },
                    Subject = subject,
                    ToRecipients = recipientList
                };

                // TODO: Call Microsoft Graph to send an email and save a copy in the Sent Items folder.
                await graphClient.Me.SendMail(email, true).Request().PostAsync();
            }
            catch (Exception e)
            {
                throw new Exception("We could not send the message: " + e.Message);
            }
        }
    }
}
