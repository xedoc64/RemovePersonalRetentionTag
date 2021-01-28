// RemovePersonalRetentionTag
//
// Author: Torsten Schlopsnies
//
// Published under MIT license

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using Serilog;

namespace RemovePersonalRetentionTag
{
    internal static class Program
    {
        

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

                // We need to start the logger here
                Log.Logger = new LoggerConfiguration()
                    .ReadFrom.AppSettings()
                    .CreateLogger();

                // Parsing the arguments and basic plausibility checks
                Log.Debug("Program started");

                Log.Debug("Parsing arguments");
                Log.Debug("mailbox: {Mailbox}", arguments.Mailbox);
                Log.Debug("logonly: {IsLogonlySet}", arguments.LogOnly);
                Log.Debug("impersonate: {IsImpersonateSet}", arguments.Impersonate);
                Log.Debug("allowredirection: {IsAllowredirectionSet}", arguments.AllowRedirection);
                Log.Debug("archive: {IsArchiveSet}", arguments.Archive);
                if (arguments.RetentionId != null) Log.Debug("Retention id filter: {RetentionId}", arguments.RetentionId);
                if (arguments.Foldername != null)
                {
                    Log.Debug("foldername: {Foldername}", arguments.RetentionId);
                }

                if (arguments.User != null)
                {
                    Log.Debug("user: {User}", arguments.User);
                }

                if (arguments.Password != null)
                {
                    Log.Debug("password: is set, will not be logged");
                }

                if (arguments.IgnoreCertificate)
                {
                    Log.Warning("Ignoring SSL error because option -ignorecertificate is set");
                    ServicePointManager.ServerCertificateValidationCallback +=
                        (sender, cert, chain, sslPolicyErrors) => true;
                }

                if (arguments.URL != null)
                {
                    Log.Debug("Server URL: {ServerUrl}", arguments.URL);
                }
                else
                {
                    Log.Debug("Server URL: using autodiscover");
                }

                // We set here the root folder from where we will searching
                if (arguments.Archive)
                {
                    Log.Debug("archive: true");
                    Log.Debug("Searching in archive instead of mailbox");
                    rootFolder = WellKnownFolderName.ArchiveMsgFolderRoot;
                }
                else
                {
                    rootFolder = WellKnownFolderName.MsgFolderRoot;
                }

                if ((arguments.Mailbox == null) || (arguments.Mailbox.Length == 0))
                {
                    Log.Error("No mailbox given. Use -help to refer to the usage");
                    Log.Error("Program stopped with failures");
                    Console.WriteLine("No mailbox given. Use -help to refer to the usage.");
                    Environment.Exit(1);
                }

                // Create the  exchange service
                ExchangeService exService;

                // connect to the server
                if (arguments.URL != null)
                {
                    // Autodiscover
                    exService = new ExService().Service(arguments.Mailbox, arguments.URL, arguments.User,
                        arguments.Password, arguments.Impersonate);
                }
                else
                {
                    exService = new ExService().Service(arguments.Mailbox, arguments.AllowRedirection, arguments.User,
                        arguments.Password, arguments.Impersonate);
                }

                if (exService == null)
                {
                    Log.Error("Error on creating the ews connection. Please check the parameters and permissions");
                    Log.Error("Program stopped with failures");
                    Environment.Exit(2);
                }

                Log.Debug("Service created");

                List<Folder> folderList = FolderUtils.Folders(exService, new FolderId(rootFolder, arguments.Mailbox));

                // We will filter the complete list, if the parameter foldername was set
                if (!string.IsNullOrEmpty(arguments.Foldername))
                {
                    Log.Information(
                        "Filtering the folder list because \"-foldername {Foldername}\" was set",
                        arguments.Foldername);

                    for (int i = folderList.Count - 1; i >= 0; i--) // yes, we need to it this way...
                    {
                        try
                        {
                            var folderPath = FolderUtils.GetFolderPath(exService, folderList[i].Id);

                            if (!(folderPath.Contains(arguments.Foldername)))
                            {
                                Log.Debug(
                                        "The folder: \"{FolderPath}\" does not match with the filter: \"{Foldername}\"",
                                        folderPath, arguments.Foldername);
                                folderList.RemoveAt(i);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex, "Exception on an ews operation");
                            Log.Error("Program stopped with failures");
                            Environment.Exit(2);
                        }
                    }
                }

                // Remove the tag, retention time and retention flag
                FolderUtils.RemoveTag(folderList, exService,
                    arguments.RetentionId?.Split(',').ToList(),
                    !arguments.LogOnly);
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
    }
}