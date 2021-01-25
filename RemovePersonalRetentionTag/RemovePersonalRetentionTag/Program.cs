// RemovePersonalRetentionTag
//
// Author: Torsten Schlopsnies
//
// Published under LGPL3.0 license

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using Microsoft.Exchange.WebServices.Data;

// Configure log4net using the .config file
[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace RemovePersonalRetentionTag
{
    class Program
    {
        // Logger       
        private static readonly log4net.ILog Log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private static FindFoldersResults _findFolders;

        /// <summary>
        /// The main function from the program
        /// </summary>
        /// <param name="args">String array containing the arguments passed to the program</param>
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                var arguments = new UtilityArguments.UtilityArguments(args);
                WellKnownFolderName rootFolder;

                if (arguments.Help)
                {
                    DisplayHelp();
                    Environment.Exit(0);
                }

                if (arguments.Usage != null)
                {
                    DisplayUsage(arguments.Usage);
                    Environment.Exit(0);
                }

                // Parsing the arguments and basic plausibility checks
                if (Log.IsInfoEnabled) Log.Info("Program started.");

                if (Log.IsDebugEnabled)
                {
                    Log.Debug("Parsing arguments");
                    Log.Debug($"mailbox: {arguments.Mailbox}");
                    Log.Debug($"logonly: {arguments.LogOnly}");
                    Log.Debug($"impersonate: {arguments.Impersonate}");
                    Log.Debug($"allowredirection: {arguments.AllowRedirection}");
                    Log.Debug($"archive: {arguments.Archive}");
                    if (arguments.RetentionId != null) Log.Debug($"Retention id filter: {arguments.RetentionId}");

                    if (arguments.Foldername != null)
                    {
                        Log.Debug($"foldername: {arguments.Foldername}");
                    }

                    if (arguments.User != null)
                    {
                        Log.Debug($"user: {arguments.User}");
                    }

                    if (arguments.Password != null)
                    {
                        Log.Debug("password: is set, will not be logged");
                    }

                    if (arguments.IgnoreCertificate)
                    {
                        Log.Warn("Ignoring SSL error because option -ignorecertificate is set");
                        ServicePointManager.ServerCertificateValidationCallback +=
                            (sender, cert, chain, sslPolicyErrors) => true;
                    }

                    if (arguments.URL != null)
                    {
                        Log.Debug($"Server URL: {arguments.URL}");
                    }
                    else
                    {
                        Log.Debug("Server URL: using autodiscover");
                    }
                }

                // We set here the root folder from where we will searching
                if (arguments.Archive)
                {
                    if (Log.IsDebugEnabled)
                    {
                        Log.Debug("archive: true");
                        Log.Debug("Searching in archive instead of mailbox.");
                    }

                    rootFolder = WellKnownFolderName.ArchiveMsgFolderRoot;
                }
                else
                {
                    if (Log.IsDebugEnabled) Log.Debug("archive: true");
                    rootFolder = WellKnownFolderName.MsgFolderRoot;
                }

                if ((arguments.Mailbox == null) || (arguments.Mailbox.Length == 0))
                {
                    Log.Error("No mailbox given. Use -help to refer to the usage.");
                    Log.Error("Program stopped with failures.");
                    Console.WriteLine("No mailbox given. Use -help to refer to the usage.");
                    Environment.Exit(1);
                }

                // Create the  exchange service
                ExchangeService exService;

                // connect to the server
                if (arguments.URL != null)
                {
                    // Autodiscover
                    exService = ConnectToExchange(arguments.Mailbox, arguments.URL, arguments.User, arguments.Password,
                        arguments.Impersonate);
                }
                else
                {
                    exService = ConnectToExchange(arguments.Mailbox, arguments.AllowRedirection, arguments.User,
                        arguments.Password, arguments.Impersonate);
                }

                if (exService == null)
                {
                    Log.Error("Error on creating the ews connection. Please check the parameters and permissions.");
                    Log.Error("Program stopped with failures.");
                    Environment.Exit(2);
                }

                if (Log.IsDebugEnabled) Log.Debug("Service created.");

                List<Folder> folderList = Folders(exService, new FolderId(rootFolder, arguments.Mailbox));

                // We will filter the complete list, if the parameter foldername was set
                if (!string.IsNullOrEmpty(arguments.Foldername))
                {
                    if (Log.IsInfoEnabled)
                        Log.Info(
                            $"Filtering the folder list because \"-foldername {arguments.Foldername}\" was set");

                    for (int i = folderList.Count - 1; i >= 0; i--) // yes, we need to it this way...
                    {
                        try
                        {
                            var folderPath = GetFolderPath(exService, folderList[i].Id);

                            if (!(folderPath.Contains(arguments.Foldername)))
                            {
                                if (Log.IsDebugEnabled)
                                    Log.Debug(
                                        $"The folder: \"{folderPath}\" does not match with the filter: \"{arguments.Foldername}\"");
                                folderList.RemoveAt(i);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Error($"Exception on an ews operation. Text from the exception: {ex}.");
                            Log.Error("Program stopped with failures.");
                            Environment.Exit(2);
                        }
                    }
                }

                // Remove the tag, retention time and retention flag
                if (arguments.RetentionId != null)
                {
                    RemoveTag(folderList, exService, arguments.RetentionId.Split(',').ToList(), !arguments.LogOnly);
                }
                else
                {
                    RemoveTag(folderList, exService, null, !arguments.LogOnly);
                }
            }
        }

        private static void DisplayUsage(string argumentsUsage)
        {
            switch (argumentsUsage)
            {
                case "mailbox":
                    Console.WriteLine("Parameter: mailbox");
                    Console.WriteLine("Parameter is mandatory");
                    Console.WriteLine("\"mailbox\" is the mailbox you would like to alter.");
                    Console.WriteLine("The program expect the primary smtp address here.");
                    break;
                case "logonly":
                    Console.WriteLine("Parameter: logonly");
                    Console.WriteLine("Parameter is optional");
                    Console.WriteLine("Only logs the folder, which have an archive or policy tag.");
                    Console.WriteLine(
                        "Depending on the settings \"RemovePersonalRetentionTag.exe.config\" we log into a file");
                    Console.WriteLine("or/and to console.");
                    break;
                case "foldername":
                    Console.WriteLine("Parameter: foldername");
                    Console.WriteLine("Parameter is optional");
                    Console.WriteLine(
                        "The program will filter the folder path (like \"Inbox\\Invoices\" depending on the foldername.");
                    Console.WriteLine(
                        "It uses \"contains\" to filter the list. That signifies \"-foldername \"inbox\" would be");
                    Console.WriteLine(
                        "\"Inbox\" including the subfolders. If you like to filter for an specific folder submit");
                    Console.WriteLine("the complete folder path");
                    break;
                case "ignorecertificate":
                    Console.WriteLine("Parameter: ignorecertificate");
                    Console.WriteLine("Parameter is optional");
                    Console.WriteLine("This ignores certificate errors when connecting to the EWS endpoint.");
                    Console.WriteLine("Usally you would like to use this together with \"-url\"");
                    break;
                case "url":
                    Console.WriteLine("Parameter: url");
                    Console.WriteLine("Parameter is optional");
                    Console.WriteLine(
                        "If you can't use autodiscover for any reasons (in an testlab for example) you can specify");
                    Console.WriteLine("the EWS endpoint here. Usally its \"https://server/EWS/Exchange.asmx\"");
                    Console.WriteLine("This won't work with Exchange Online.");
                    Console.WriteLine("Whenever it's possible use autodiscover.");
                    break;
                case "user":
                    Console.WriteLine("Parameter: user");
                    Console.WriteLine("Parameter is optional");
                    Console.WriteLine(
                        "Specify the username (primary smtp address) which will be used for altering the mailbox");
                    Console.WriteLine(
                        "If you specify a user you need also to specify the password. If no user is specified");
                    Console.WriteLine(
                        "the credentials from the current user which is running the program will be used.");
                    break;
                case "password":
                    Console.WriteLine("Parameter: user");
                    Console.WriteLine("Parameter is optional. It's mandatory when \"-user\" parameter is set.");
                    Console.WriteLine("Specify the password for \"user\".");
                    Console.WriteLine("Used together with the option \"-user\".");
                    break;
                case "impersonate":
                    Console.WriteLine("Parameter: impersonate");
                    Console.WriteLine("Parameter is optional");
                    Console.WriteLine("Parameter will be used, when you want to alter a mailbox other than yours.");
                    Console.WriteLine("You need ApplicationImpersonation rights on the exchange server.");
                    break;
                case "retentionid":
                    Console.WriteLine("Parameter: retentionid");
                    Console.WriteLine("Parameter is optional");
                    Console.WriteLine("You can use this to filter the retention id which will be removed.");
                    Console.WriteLine("This can be used when the user have more than one personal policy tag applied.");
                    Console.WriteLine(
                        "You can get the retention id in the Exchange Server powershell from the retention policy tags");
                    break;
                case "archive":
                    Console.WriteLine("Parameter: archive");
                    Console.WriteLine("Parameter is optional");
                    Console.WriteLine("Search folders inside the online archive instead of the mailbox.");
                    break;
                default:
                    DisplayHelp();
                    break;
            }
        }

        private static void RemoveTag(List<Folder> folderList, ExchangeService exService, List<string> retentionId,
            bool removeTag)
        {
            foreach (var folder in folderList)
            {
                var folderChanged = false;
                var oFolder = Folder.Bind(exService, folder.Id);
                if (oFolder.ArchiveTag != null)
                {
                    if (Log.IsInfoEnabled)
                    {
                        Log.Info($"Folder with archive tag found, ID: {folder.Id}");
                        Log.Info($"Folder name: {folder.DisplayName}");
                        Log.Info($"Folder path: {GetFolderPath(exService, folder.Id)}");
                        Log.Info($"Retention id: {oFolder.ArchiveTag.RetentionId}");
                    }

                    if ((retentionId != null) && (retentionId.Contains(oFolder.ArchiveTag.RetentionId.ToString())))
                    {
                        if (removeTag)
                        {
                            if (Log.IsInfoEnabled) Log.Info("Removing the archive tag.");
                            try
                            {
                                oFolder.ArchiveTag = null;
                                folderChanged = true;
                            }
                            catch (Exception e)
                            {
                                Log.Error(
                                    $"Error on removing archive tag from folder: {folder.Id}. Path: {GetFolderPath(exService, folder.Id)}");
                                Log.Error($"Exception: {e}");
                            }
                        }
                    }
                    else if (removeTag)
                    {
                        if (Log.IsInfoEnabled) Log.Info("Removing the archive tag.");
                        try
                        {
                            oFolder.ArchiveTag = null;
                            folderChanged = true;
                        }
                        catch (Exception e)
                        {
                            Log.Error(
                                $"Error on removing archive tag from folder: {folder.Id}. Path: {GetFolderPath(exService, folder.Id)}");
                            Log.Error($"Exception: {e}");
                        }
                    }
                }

                if (oFolder.PolicyTag != null)
                {
                    if (Log.IsInfoEnabled)
                    {
                        Log.Info($"Folder with policy tag found, ID: {folder.Id}");
                        Log.Info($"Folder name: {folder.DisplayName}");
                        Log.Info($"Folder path: {GetFolderPath(exService, folder.Id)}");
                        Log.Info($"Retention id: {oFolder.PolicyTag.RetentionId}");
                    }

                    if ((retentionId != null) && (retentionId.Contains(oFolder.PolicyTag.RetentionId.ToString())))
                    {
                        if (removeTag)
                        {
                            if (Log.IsInfoEnabled) Log.Info("Removing the policy tag.");
                            try
                            {
                                oFolder.PolicyTag = null;
                                folderChanged = true;
                            }
                            catch (Exception e)
                            {
                                Log.Error(
                                    $"Error on removing policy tag from folder: {folder.Id}. Path: {GetFolderPath(exService, folder.Id)}");
                                Log.Error($"Exception: {e}");
                            }
                        }
                    }
                    else if (removeTag)
                    {
                        if (Log.IsInfoEnabled) Log.Info("Removing the policy tag.");
                        try
                        {
                            oFolder.PolicyTag = null;
                            folderChanged = true;
                        }
                        catch (Exception e)
                        {
                            Log.Error(
                                $"Error on removing policy tag from folder: {folder.Id}. Path: {GetFolderPath(exService, folder.Id)}");
                            Log.Error($"Exception: {e}");
                        }
                    }
                }

                if (folderChanged)
                {
                    try
                    {
                        oFolder.Update();
                        if (Log.IsInfoEnabled) Log.Info("Tag removed successfully.");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                        throw;
                    }
                }
            }
        }

        /// <summary>
        /// Display a basic help
        /// </summary>
        private static void DisplayHelp()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine(
                "RemovePersonalRetentionTag.exe -mailbox \"user@example.com\" [-logonly] [-foldername \"Inbox\"]  [-ignorecertificate] [-url \"https://server/EWS/Exchange.asmx\"] [-user \"user@example.com\"] [-password \"Pa$$w0rd\"] [-impersonate] [-retentionid \"a7966968-dadf-4df7-ae87-4482686b4634\" [-archive]");
            Console.WriteLine(
                "For more information use the parameter -usage with a parameter you would like to know about more. E.g. -usage \"mailbox\"");
        }

        /// <summary>
        /// Connect to Exchange using AutoDiscover for the given email address
        /// </summary>
        /// <param name="mailboxId">The users email address</param>
        /// <param name="allowredirection"></param>
        /// <param name="user"></param>
        /// <param name="password"></param>
        /// <param name="impersonation"></param>
        /// <returns>Exchange Web Service binding</returns>
        private static ExchangeService ConnectToExchange(string mailboxId, bool allowredirection, string user,
            string password, bool impersonation)
        {
            if (Log.IsInfoEnabled) Log.Info($"Connect to mailbox {mailboxId}");
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
            catch (Exception ex)
            {
                Log.Error("Connection to mailbox failed", ex);
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
        private static ExchangeService ConnectToExchange(string mailboxId, string url, string user, string password,
            bool impersonation)
        {
            if (Log.IsInfoEnabled) Log.Info($"Connect to mailbox {mailboxId}");
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
            catch (Exception ex)
            {
                Log.Error("Connection to mailbox failed", ex);
            }

            return null;
        }

        /// <summary>
        /// Get a single mailbox folder path
        /// </summary>
        /// <param name="service">The active EWs connection</param>
        /// <param name="id">The mailbox folder Id</param>
        /// <returns>A string containing the current mailbox folder path</returns>
        private static string GetFolderPath(ExchangeService service, FolderId id)
        {
            try
            {
                var folderPathProperty = new ExtendedPropertyDefinition(0x66B5, MapiPropertyType.String);

                var psset1 = new PropertySet(BasePropertySet.FirstClassProperties);
                psset1.Add(folderPathProperty);

                var folderwithPath = Folder.Bind(service, id, psset1);

                if (folderwithPath.TryGetProperty(folderPathProperty, out var folderPathVal))
                {
                    // because the FolderPath contains characters we don't want, we need to fix it
                    var folderPathTemp = folderPathVal.ToString();
                    if (folderPathTemp.Contains("￾"))
                    {
                        return folderPathTemp.Replace("￾", "\\");
                    }
                    else
                    {
                        return folderPathTemp;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("Failed to get folder path", ex);
            }

            return "";
        }

        /// <summary>
        /// Find all folders under MsgRootFolder
        /// </summary>
        /// <param name="service"></param>
        /// <param name="searchRootFolder"></param>
        /// <returns>Result of a folder search operation</returns>
        private static List<Folder> Folders(ExchangeService service, FolderId searchRootFolder)
        {
            // try to find all folder that are unter MsgRootFolder
            int pageSize = 100;
            int pageOffset = 0;
            bool moreItems = true;
            var view = new FolderView(pageSize, pageOffset);
            var resultFolders = new List<Folder>();


            var propertySet =
                new PropertySet(BasePropertySet.FirstClassProperties, FolderSchema.DisplayName);

            view.PropertySet = propertySet;
            view.Traversal = FolderTraversal.Deep;

            while (moreItems)
            {
                try
                {
                    _findFolders = service.FindFolders(searchRootFolder, view);

                    moreItems = _findFolders.MoreAvailable;

                    foreach (var folder in _findFolders)
                    {
                        resultFolders.Add(folder);
                    }

                    // if we have more folders than we have to page
                    if (moreItems) view.Offset += pageSize;
                }
                catch (Exception ex)
                {
                    Log.Error("Failed to fetch folders.", ex);
                    Log.Error("Program ended with errors.");
                    moreItems = false;
                    Environment.Exit(2);
                }
            }

            return resultFolders;
        }

        /// <summary>
        /// Redirection handler if -allowredirection is set
        /// </summary>
        /// <param name="url">The url which the program will connect to.</param>
        /// <returns></returns>
        public static bool RedirectionCallback(string url)
        {
            // Return true if the URL is an HTTPS URL.
            return url.ToLower().StartsWith("https://");
        }
    }
}