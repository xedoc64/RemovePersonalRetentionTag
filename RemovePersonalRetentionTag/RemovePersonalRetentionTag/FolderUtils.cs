using System;
using System.Collections.Generic;
using Serilog;
using Microsoft.Exchange.WebServices.Data;

namespace RemovePersonalRetentionTag
{
    public static class FolderUtils
    {
        private static FindFoldersResults _findFolders;
        
        public static void RemoveTag(List<Folder> folderList, ExchangeService exService, List<string> retentionId,
            bool removeTag)
        {
            int found = 0;
            int removed = 0;

            foreach (var folder in folderList)
            {
                var folderChanged = false;
                var oFolder = Folder.Bind(exService, folder.Id);
                if (oFolder.ArchiveTag != null)
                {
                    found++;
                    Log.Information("Folder with archive tag found, ID: {FolderID}", folder.Id);
                    Log.Information("Folder name: {FolderDisplayName}", folder.DisplayName);
                    Log.Information("Folder path: {FolderPath}", GetFolderPath(exService, folder.Id));
                    Log.Information("Retention id: {FolderRetentionId}", oFolder.ArchiveTag.RetentionId);
                    
                    if ((retentionId != null) && (retentionId.Contains(oFolder.ArchiveTag.RetentionId.ToString())))
                    {
                        if (removeTag)
                        {
                            Log.Information("Removing the archive tag");
                            try
                            {
                                oFolder.ArchiveTag = null;
                                folderChanged = true;
                            }
                            catch (Exception e)
                            {
                                Log.Error(
                                    "Error on removing archive tag from folder: " +
                                    "{FolderId}. Path: {FolderPath}", 
                                    folder.Id, GetFolderPath(exService, folder.Id));
                                Log.Error(e, "Exception:");
                            }
                        }
                    }
                    else if (removeTag)
                    {
                        Log.Information("Removing the archive tag");
                        try
                        {
                            oFolder.ArchiveTag = null;
                            folderChanged = true;
                            removed++;
                        }
                        catch (Exception e)
                        {
                            Log.Error(
                                "Error on removing archive tag from folder: {FolderId}. " +
                                "Path: {FolderPath}",folder.Id, 
                                GetFolderPath(exService, folder.Id));
                            Log.Error(e, "Exception:");
                        }
                    }
                }

                if (oFolder.PolicyTag != null)
                {
                     Log.Information("Folder with policy tag found, ID: {FolderID}", folder.Id);
                     Log.Information("Folder name: {FolderDisplayName}",folder.DisplayName);
                     Log.Information("Folder path: {FolderPath}", GetFolderPath(exService, folder.Id));
                     Log.Information("Retention id: {FolderRetentionId}", oFolder.PolicyTag.RetentionId);
                    

                    if ((retentionId != null) && (retentionId.Contains(oFolder.PolicyTag.RetentionId.ToString())))
                    {
                        if (removeTag)
                        {
                            Log.Information("Removing the policy tag");
                            try
                            {
                                oFolder.PolicyTag = null;
                                folderChanged = true;
                            }
                            catch (Exception e)
                            {
                                Log.Error(
                                    "Error on removing policy tag from folder: {FolderID}. " +
                                    "Path: {FolderPath}", folder.Id, GetFolderPath(exService, folder.Id));
                                Log.Error(e ,"Exception:");
                            }
                        }
                    }
                    else if (removeTag)
                    {
                        Log.Information("Removing the policy tag");
                        try
                        {
                            oFolder.PolicyTag = null;
                            folderChanged = true;
                        }
                        catch (Exception e)
                        {
                            Log.Error(
                                "Error on removing policy tag from folder: {FolderID}. Path: {FolderPath}",
                                folder.Id, GetFolderPath(exService, folder.Id));
                            Log.Error(e ,"Exception:");
                        }
                    }
                }

                if (folderChanged)
                {
                    try
                    {
                        oFolder.Update();
                        Log.Information("Tag removed successfully");
                        
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                        throw;
                    }
                }
            }
            Log.Information("Folders with a personal retention tag found: {found}");
            Log.Information("Folders with a personal retention tag removed: {removed}");
        }
        
                /// <summary>
        /// Get a single mailbox folder path
        /// </summary>
        /// <param name="service">The active EWs connection</param>
        /// <param name="id">The mailbox folder Id</param>
        /// <returns>A string containing the current mailbox folder path</returns>
        public static string GetFolderPath(ExchangeService service, FolderId id)
        {
            try
            {
                var folderPathProperty = new ExtendedPropertyDefinition(0x66B5, MapiPropertyType.String);

                var propertySet = new PropertySet(BasePropertySet.FirstClassProperties) {folderPathProperty};

                var folderIncludingPath = Folder.Bind(service, id, propertySet);

                if (folderIncludingPath.TryGetProperty(folderPathProperty, out var folderPathVal))
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
            catch (Exception e)
            {
                Log.Error(e,"Failed to get folder path");
            }

            return "";
        }

        /// <summary>
        /// Find all folders under MsgRootFolder
        /// </summary>
        /// <param name="service"></param>
        /// <param name="searchRootFolder"></param>
        /// <returns>Result of a folder search operation</returns>
        public static List<Folder> Folders(ExchangeService service, FolderId searchRootFolder)
        {
            // try to find all folder that are under MsgRootFolder
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
                catch (Exception e)
                {
                    Log.Error(e, "Failed to fetch folders");
                    Log.Error("Program ended with errors");
                    moreItems = false;
                    Environment.Exit(2);
                }
            }

            return resultFolders;
        }
    }
}