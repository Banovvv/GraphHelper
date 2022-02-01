using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GHelper
{
    public class GHelper
    {
        public static async Task<MailFolder> GetMailFolderByName(GraphServiceClient graphClient, string user, string folderName)
        {
            var graphFolder = await graphClient.Users[user].MailFolders.Inbox.ChildFolders
                                               .Request()
                                               .Filter($"startsWith(displayName, '{folderName}')")
                                               .GetAsync();
            return graphFolder.First();
        }

        public static async Task<MailFolder> GetMailFolderByID(GraphServiceClient graphClient, string user, string folderID)
        {
            var graphFolder = await graphClient.Users[user].MailFolders[folderID]
                                               .Request()
                                               .GetAsync();
            return graphFolder;
        }

        public static async Task<string> GetMailFolderID(GraphServiceClient graphClient, string user, string folderName)
        {
            var graphFolder = await graphClient.Users[user].MailFolders.Inbox.ChildFolders
                                               .Request()
                                               .Filter($"startsWith(displayName, '{folderName}')")
                                               .GetAsync();
            string folderID = graphFolder.First().Id;
            return folderID;
        }

        public static async Task<IMailFolderMessagesCollectionPage> GetInboxEmails(GraphServiceClient graphClient, string user)
        {
            var inboxMessages = await graphClient.Users[user].MailFolders.Inbox.Messages
                                                 .Request()
                                                 .Top(50)
                                                 .GetAsync();
            return inboxMessages;
        }

        public static async Task<IMailFolderMessagesCollectionPage> GetInboxEmails(GraphServiceClient graphClient, string user, int numberOfEmails)
        {
            var inboxMessages = await graphClient.Users[user].MailFolders.Inbox.Messages
                                                 .Request()
                                                 .Top(numberOfEmails)
                                                 .GetAsync();
            return inboxMessages;
        }

        public static async Task<IMailFolderMessagesCollectionPage> GetSpecificFolderEmails(GraphServiceClient graphClient, string user, string folderName)
        {
            var graphFolder = await graphClient.Users[user].MailFolders.Inbox.ChildFolders
                                               .Request()
                                               .Filter($"startsWith(displayName, '{folderName}')")
                                               .GetAsync();
            var emailFolderID = graphFolder.First().Id;

            var folderMessages = await graphClient.Users[user].MailFolders[emailFolderID].Messages
                                                 .Request()
                                                 .Top(50)
                                                 .GetAsync();
            return folderMessages;
        }

        public static async Task SendExceptionEmail(GraphServiceClient graphClient, string user, string recipient, Exception ex)
        {
            var message = new Message
            {
                Subject = $"{ex.Message}",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = ex.StackTrace
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipient
                        }
                    }
                }
            };

            await graphClient.Users[user]
                .SendMail(message, true)
                .Request()
                .PostAsync();
        }

        public static GraphServiceClient GetGraphClient(string tenantID, string clientID, string secret)
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            ClientSecretCredential clientSecretCredential = new ClientSecretCredential(
                tenantID, clientID, secret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            return graphClient;
        }
    }
}
