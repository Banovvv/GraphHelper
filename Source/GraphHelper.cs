using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GraphHelper
{
    public class GraphHelper
    {
        /// <summary>
        /// This method returns a <see cref="Microsoft.Graph.MailFolder"/> by given a <see cref="string"/> <paramref name="folderName"/>
        /// </summary>
        /// <param name="graphClient">An instance of a Microsoft Graph Client</param>
        /// <param name="user">A string reperesenting the email of a Microsoft Graph user </param>
        /// <param name="folderName">A string reprsenting the requested folder name</param>
        /// <returns>The <see cref="Microsoft.Graph.MailFolder"/> with the requested name</returns>
        public static async Task<MailFolder> GetMailFolderByName(GraphServiceClient graphClient, string user, string folderName)
        {
            var graphFolder = await graphClient.Users[user].MailFolders.Inbox.ChildFolders
                                               .Request()
                                               .Filter($"startsWith(displayName, '{folderName}')")
                                               .GetAsync();
            return graphFolder.First();
        }

        /// <summary>
        /// This method returns a <see cref="Microsoft.Graph.MailFolder"/> by given a <see cref="string"/> <paramref name="folderID"/>
        /// </summary>
        /// <param name="graphClient">An instance of a Microsoft Graph Client</param>
        /// <param name="user">A string reperesenting the email of a Microsoft Graph user </param>
        /// <param name="folderID">A string reprsenting the requested folder's ID</param>
        /// <returns>The <see cref="Microsoft.Graph.MailFolder"/> with the requested Folder ID</returns>
        public static async Task<MailFolder> GetMailFolderByID(GraphServiceClient graphClient, string user, string folderID)
        {
            var graphFolder = await graphClient.Users[user].MailFolders[folderID]
                                               .Request()
                                               .GetAsync();
            return graphFolder;
        }
        /// <summary>
        /// This method returns a <see cref="string"/> representig the ID of the <see cref="Microsoft.Graph.MailFolder"/> by given a <see cref="string"/> <paramref name="folderName"/>
        /// </summary>
        /// <param name="graphClient">An instance of a Microsoft Graph Client</param>
        /// <param name="user">A string reperesenting the email of a Microsoft Graph user</param>
        /// <param name="folderName">A string reprsenting the requested folder name</param>
        /// <returns>A <see cref="string"/> representig the ID of the <see cref="Microsoft.Graph.MailFolder"/></returns>
        public static async Task<string> GetMailFolderID(GraphServiceClient graphClient, string user, string folderName)
        {
            var graphFolder = await graphClient.Users[user].MailFolders.Inbox.ChildFolders
                                               .Request()
                                               .Filter($"startsWith(displayName, '{folderName}')")
                                               .GetAsync();
            string folderID = graphFolder.First().Id;
            return folderID;
        }

        /// <summary>
        /// This method return the first 50 emails of the user's Inbox as a <see cref="Microsoft.Graph.IMailFolderMessagesCollectionPage"/>
        /// </summary>
        /// <param name="graphClient">An instance of a Microsoft Graph Client</param>
        /// <param name="user">A string reperesenting the email of a Microsoft Graph user</param>
        /// <returns>A <see cref="Microsoft.Graph.IMailFolderMessagesCollectionPage"/> containing the first 50 emails of the user's Inbox</returns>
        public static async Task<IMailFolderMessagesCollectionPage> GetInboxEmails(GraphServiceClient graphClient, string user)
        {
            var inboxMessages = await graphClient.Users[user].MailFolders.Inbox.Messages
                                                 .Request()
                                                 .Top(50)
                                                 .GetAsync();
            return inboxMessages;
        }

        /// <summary>
        /// This method return the first <see cref="int"/> <paramref name="numberOfEmails"/> emails of the user's Inbox as a <see cref="Microsoft.Graph.IMailFolderMessagesCollectionPage"/>
        /// </summary>
        /// <param name="graphClient">An instance of a Microsoft Graph Client</param>
        /// <param name="user">A string reperesenting the email of a Microsoft Graph user</param>
        /// <param name="numberOfEmails">An integer representing the number of requested emails</param>
        /// <returns>A <see cref="Microsoft.Graph.IMailFolderMessagesCollectionPage"/> containing the first <see cref="int"/> <paramref name="numberOfEmails"/> emails of the user's Inbox</returns>
        public static async Task<IMailFolderMessagesCollectionPage> GetInboxEmails(GraphServiceClient graphClient, string user, int numberOfEmails)
        {
            var inboxMessages = await graphClient.Users[user].MailFolders.Inbox.Messages
                                                 .Request()
                                                 .Top(numberOfEmails)
                                                 .GetAsync();
            return inboxMessages;
        }

        /// <summary>
        /// This method return the first 50 emails of the user's <see cref="string"/> <paramref name="folderName"/> as a <see cref="Microsoft.Graph.IMailFolderMessagesCollectionPage"/>
        /// </summary>
        /// <param name="graphClient">An instance of a Microsoft Graph Client</param>
        /// <param name="user">A string reperesenting the email of a Microsoft Graph user</param>
        /// <param name="folderName">A string reprsenting the requested folder name</param>
        /// <returns>A <see cref="Microsoft.Graph.IMailFolderMessagesCollectionPage"/> containing the first 50 emails of the user's <see cref="string"/> <paramref name="folderName"/></returns>
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

        /// <summary>
        /// This method send an email to a <see cref="string"/> <paramref name="recipient"/> containing the stack trace of an <see cref="Exception"/> <paramref name="ex"/> as an email body and the <see cref="Exception"/>'s message as an email subject
        /// </summary>
        /// <param name="graphClient">An instance of a Microsoft Graph Client</param>
        /// <param name="user">A string reperesenting the email of a Microsoft Graph user</param>
        /// <param name="recipient">A string representing the recepient's email address</param>
        /// <param name="ex">An instance of the Exception</param>
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

        /// <summary>
        /// This method return an instance of <see cref="Microsoft.Graph.GraphServiceClient"/> for a given <paramref name="tenantID"/>, <paramref name="clientID"/> and <paramref name="secret"/>
        /// </summary>
        /// <param name="tenantID">A string representing the Tenant's ID</param>
        /// <param name="clientID">A string representing the Clients's ID</param>
        /// <param name="secret">A string representing the Client's Secret</param>
        /// <returns>An instance of <see cref="Microsoft.Graph.GraphServiceClient"/> for a given <paramref name="tenantID"/>, <paramref name="clientID"/> and <paramref name="secret"/></returns>
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
