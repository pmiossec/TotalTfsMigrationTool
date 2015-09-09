using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.TeamFoundation.Proxy;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Xml;
using System.IO;
using log4net.Config;
using log4net;

namespace TFSProjectMigration
{
    public class WorkItemRead
    {

        TfsTeamProjectCollection tfs;
        public WorkItemStore store;
        public QueryHierarchy queryCol;
        String projectName;
        public WorkItemTypeCollection workItemTypes;
        private static readonly ILog logger = LogManager.GetLogger(typeof(TFSWorkItemMigrationUI));

        public WorkItemRead(TfsTeamProjectCollection tfs, Project sourceProject)
        {
            this.tfs = tfs;
            projectName = sourceProject.Name;
            store = (WorkItemStore)tfs.GetService(typeof(WorkItemStore));
            queryCol = store.Projects[sourceProject.Name].QueryHierarchy;
            workItemTypes = store.Projects[sourceProject.Name].WorkItemTypes;
        }

        /* Get required work items from project and save existing attachments of workitems to local folder */
        public WorkItemCollection GetWorkItems(string project)
        {
            WorkItemCollection workItemCollection = store.Query(" SELECT * " +
                                                                 " FROM WorkItems " +
                                                                 " WHERE [System.TeamProject] = '" + project +
                                                                 "'AND [System.State] <> 'Closed' ORDER BY [System.Id]");
            SaveAttachments(workItemCollection);
            return workItemCollection;
        }


        public WorkItemCollection GetWorkItems(string project, bool IsNotIncludeClosed, bool IsNotIncludeRemoved)
        {
            String query = "";
            if (IsNotIncludeClosed && IsNotIncludeRemoved)
            {
                query = String.Format(" SELECT * " +
                                                    " FROM WorkItems " +
                                                    " WHERE [System.TeamProject] = '" + project +
                                                    "'AND [System.State] <> 'Closed' AND [System.State] <> 'Removed' ORDER BY [System.Id]");
            }

            else if (IsNotIncludeRemoved)
            {
                query = String.Format(" SELECT * " +
                                                   " FROM WorkItems " +
                                                   " WHERE [System.TeamProject] = '" + project +
                                                   "'AND [System.State] <> 'Removed' ORDER BY [System.Id]");
            }
            else if (IsNotIncludeClosed)
            {
                query = String.Format(" SELECT * " +
                                                   " FROM WorkItems " +
                                                   " WHERE [System.TeamProject] = '" + project +
                                                   "'AND [System.State] <> 'Closed'  ORDER BY [System.Id]");
            }
            else
            {
                query = String.Format(" SELECT * " +
                                                   " FROM WorkItems " +
                                                   " WHERE [System.TeamProject] = '" + project +
                                                   "' ORDER BY [System.Id]");
            }
            System.Diagnostics.Debug.WriteLine(query);
            WorkItemCollection workItemCollection = store.Query(query);
            SaveAttachments(workItemCollection);
            return workItemCollection;
        }
        /* Save existing attachments of workitems to local folders of workitem ID */
        private void SaveAttachments(WorkItemCollection workItemCollection)
        {
            if (!Directory.Exists(@"Attachments"))
            {
                Directory.CreateDirectory(@"Attachments");
            }
            else
            {
                EmptyFolder(new DirectoryInfo(@"Attachments"));
            }

            System.Net.WebClient webClient = new System.Net.WebClient();
            webClient.UseDefaultCredentials = true;

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
                            if (!File.Exists(path + "\\" + att.Name))
                            {
                                webClient.DownloadFile(att.Uri, path + "\\" + att.Name);
                            }
                            else 
                            {
                                webClient.DownloadFile(att.Uri, path + "\\" + att.Id + "_" + att.Name);
                            }
                           
                        }
                        catch (Exception)
                        {
                            logger.Info("Error downloading attachment for work item : " + wi.Id + " Type: " + wi.Type.Name);
                        }

                    }
                }
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
            ICommonStructureService css = (ICommonStructureService)tfs.GetService(typeof(ICommonStructureService));
            //Gets Area/Iteration base Project
            ProjectInfo projectInfo = css.GetProjectFromName(projectName);
            NodeInfo[] nodes = css.ListStructures(projectInfo.Uri);

            //GetNodes can use with:
            //Area = 0
            //Iteration = 1
            XmlElement AreaTree = css.GetNodesXml(new string[] { nodes[0].Uri }, true);
            XmlElement IterationsTree = css.GetNodesXml(new string[] { nodes[1].Uri }, true);

            XmlNode AreaNodes = AreaTree.ChildNodes[0];
            XmlNode IterationsNodes = IterationsTree.ChildNodes[0];

            return new XmlNode[] { AreaNodes, IterationsNodes };
        }


    }
}
