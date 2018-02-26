using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Controls;
using System.Xml;
using log4net;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TFSProjectMigration
{
    public class WorkItemRead
    {
        private TfsTeamProjectCollection _tfs;
        public WorkItemStore Store;
        public QueryHierarchy QueryCol;
        private String _projectName;
        public WorkItemTypeCollection WorkItemTypes;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(TfsWorkItemMigrationUi));

        public WorkItemRead(TfsTeamProjectCollection tfs, Project sourceProject)
        {
            _tfs = tfs;
            _projectName = sourceProject.Name;
            Store = (WorkItemStore)tfs.GetService(typeof(WorkItemStore));
            QueryCol = Store.Projects[sourceProject.Name].QueryHierarchy;
            WorkItemTypes = Store.Projects[sourceProject.Name].WorkItemTypes;
        }

        /* Get required work items from project and save existing attachments of workitems to local folder */
        public WorkItemCollection GetWorkItems(string project, ProgressBar progressBar)
        {
            WorkItemCollection workItemCollection = Store.Query(" SELECT * " +
                                                                 " FROM WorkItems " +
                                                                 " WHERE [System.TeamProject] = '" + project +
                                                                 "' AND [System.State] <> 'Closed' ORDER BY [System.Id]");
            SaveAttachments(workItemCollection, progressBar);
            return workItemCollection;
        }

        public WorkItemCollection GetWorkItems(string project, bool isNotIncludeClosed, bool isNotIncludeRemoved, ProgressBar progressBar)
        {
            String query;
            if (isNotIncludeClosed && isNotIncludeRemoved)
            {
                query = String.Format(" SELECT * " +
                                                    " FROM WorkItems " +
                                                    " WHERE [System.TeamProject] = '" + project +
                                                    "' AND [System.State] <> 'Closed' AND [System.State] <> 'Removed' ORDER BY [System.Id]");
            }

            else if (isNotIncludeRemoved)
            {
                query = String.Format(" SELECT * " +
                                                   " FROM WorkItems " +
                                                   " WHERE [System.TeamProject] = '" + project +
                                                   "' AND [System.State] <> 'Removed' ORDER BY [System.Id]");
            }
            else if (isNotIncludeClosed)
            {
                query = String.Format(" SELECT * " +
                                                   " FROM WorkItems " +
                                                   " WHERE [System.TeamProject] = '" + project +
                                                   "' AND [System.State] <> 'Closed'  ORDER BY [System.Id]");
            }
            else
            {
                query = String.Format(" SELECT * " +
                                                   " FROM WorkItems " +
                                                   " WHERE [System.TeamProject] = '" + project +
                                                   "' ORDER BY [System.Id]");
            }
            Debug.WriteLine(query);
            WorkItemCollection workItemCollection = Store.Query(query);
            SaveAttachments(workItemCollection, progressBar);
            return workItemCollection;
        }
        /* Save existing attachments of workitems to local folders of workitem ID */
        private void SaveAttachments(WorkItemCollection workItemCollection, ProgressBar progressBar)
        {
            if (!Directory.Exists(@"Attachments"))
            {
                Directory.CreateDirectory(@"Attachments");
            }
            else
            {
                EmptyFolder(new DirectoryInfo(@"Attachments"));
            }

            WebClient webClient = new WebClient();
            webClient.UseDefaultCredentials = true;

            int index = 0;
            foreach (WorkItem wi in workItemCollection)
            {
                if (wi.AttachedFileCount > 0)
                {
                    foreach (Attachment att in wi.Attachments)
                    {
                        try
                        {
                            String path = @"Attachments\" + wi.Id;
                            bool folderExists = Directory.Exists(path);
                            if (!folderExists)
                            {
                                Directory.CreateDirectory(path);
                            }
                            var fileInfo = new FileInfo(path + "\\" + att.Name);
                            if (!fileInfo.Exists)
                            {
                                webClient.DownloadFile(att.Uri, path + "\\" + att.Name);
                            }
                            else if (fileInfo.Length != att.Length)
                            {
                                webClient.DownloadFile(att.Uri, path + "\\" + att.Id + "_" + att.Name);
                            }
                        }
                        catch (Exception)
                        {
                            Logger.Info("Error downloading attachment for work item : " + wi.Id + " Type: " + wi.Type.Name);
                        }
                    }
                }
                index++;
                var index1 = index;
                progressBar.Dispatcher.BeginInvoke(new Action(delegate
                {
                    progressBar.Value = index1 / (float)workItemCollection.Count * 100;
                }));
            }
        }


        /*Delete all subfolders and files in given folder*/
        private void EmptyFolder(DirectoryInfo directoryInfo)
        {
            foreach (FileInfo file in directoryInfo.GetFiles())
            {
                file.Delete();
            }
            foreach (DirectoryInfo subfolder in directoryInfo.GetDirectories())
            {
                EmptyFolder(subfolder);
                subfolder.Delete();
            }
        }

        /* Return Areas and Iterations of the project */
        public XmlNode[] PopulateIterations()
        {
            ICommonStructureService css = (ICommonStructureService)_tfs.GetService(typeof(ICommonStructureService));
            //Gets Area/Iteration base Project
            ProjectInfo projectInfo = css.GetProjectFromName(_projectName);
            NodeInfo[] nodes = css.ListStructures(projectInfo.Uri);

            XmlElement areaTree = css.GetNodesXml(new[] { nodes.Single(n => n.StructureType == "ProjectModelHierarchy").Uri }, true);
            XmlElement iterationsTree = css.GetNodesXml(new[] { nodes.Single(n => n.StructureType == "ProjectLifecycle").Uri }, true);

            XmlNode areaNodes = areaTree.ChildNodes[0];
            XmlNode iterationsNodes = iterationsTree.ChildNodes[0];

            return new[] { areaNodes, iterationsNodes };
        }

        




    }
}
