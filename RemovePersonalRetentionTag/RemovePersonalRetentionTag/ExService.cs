using System;
using Microsoft.Exchange.WebServices.Data;
using Serilog;

namespace RemovePersonalRetentionTag
{
    public class ExService
    {

        
        /// <summary>
        /// Connect to Exchange using AutoDiscover for the given email address
        /// </summary>
        /// <param name="mailboxId">The users email address</param>
        /// <param name="allowredirection"></param>
        /// <param name="user"></param>
        /// <param name="password"></param>
        /// <param name="impersonation"></param>
        /// <returns>Exchange Web Service binding</returns>
        public ExchangeService Service(string mailboxId, bool allowredirection, string user,
            string password, bool impersonation)
        {
            Log.Information("Connect to mailbox {Mailbox}", mailboxId);
            try
            {
                var service = new ExchangeService();

                if ((user == null) | (password == null))
                {
                    service.UseDefaultCredentials = true;
                }
                else
                {
                    service.Credentials = new WebCredentials(user, password);
                }

                if (allowredirection)
                {
                    service.AutodiscoverUrl(mailboxId, RedirectionCallback);
                }
                else
                {
                    service.AutodiscoverUrl(mailboxId);
                }

                if (impersonation)
                {
                    service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailboxId);
                }

                return service;
            }
            catch (Exception e)
            {
                Log.Error(e, "Connection to mailbox failed");
            }

            return null;
        }

        /// <summary>
        /// Connect to Exchange using AutoDiscover for the given email address
        /// </summary>
        /// <param name="mailboxId">The users email address</param>
        /// <param name="url"></param>
        /// <param name="user"></param>
        /// <param name="password"></param>
        /// <param name="impersonation"></param>
        /// <returns>Exchange Web Service binding</returns>
        public ExchangeService Service(string mailboxId, string url, string user, string password,
            bool impersonation)
        {
            Log.Information("Connect to mailbox {Mailbox}", mailboxId);
            try
            {
                var service = new ExchangeService();

                if ((user == null) | (password == null))
                {
                    service.UseDefaultCredentials = true;
                }
                else
                {
                    service.Credentials = new WebCredentials(user, password);
                }

                service.Url = new Uri(url);
                if (impersonation)
                {
                    service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailboxId);
                }

                return service;
            }
            catch (Exception e)
            {
                Log.Error(e, "Connection to mailbox failed");
            }

            return null;
        }
        
        
        /// <summary>
        /// Redirection handler if -allowredirection is set
        /// </summary>
        /// <param name="url">The url which the program will connect to.</param>
        /// <returns></returns>
        private static bool RedirectionCallback(string url)
        {
            // Return true if the URL is an HTTPS URL.
            return url.ToLower().StartsWith("https://");
        }
    }
}