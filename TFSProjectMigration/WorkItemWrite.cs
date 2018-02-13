using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Controls;
using System.Xml;
using log4net;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Proxy;

namespace TFSProjectMigration
{
    public class WorkItemWrite
    {
        private readonly TfsTeamProjectCollection _tfs;
        public WorkItemStore Store;
        public QueryHierarchy QueryCol;
        private readonly Project _destinationProject;
        private readonly String _projectName;
        public XmlNode AreaNodes;
        public XmlNode IterationsNodes;
        private readonly WorkItemTypeCollection _workItemTypes;
        public Hashtable ItemMap;
        public Hashtable ItemMapCic;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(TfsWorkItemMigrationUi));

        public WorkItemWrite(TfsTeamProjectCollection tfs, Project destinationProject)
        {
            _tfs = tfs;
            _projectName = destinationProject.Name;
            _destinationProject = destinationProject;
            Store = new WorkItemStore(tfs, WorkItemStoreFlags.BypassRules);
            QueryCol = Store.Projects[destinationProject.Name].QueryHierarchy;
            _workItemTypes = Store.Projects[destinationProject.Name].WorkItemTypes;
            ItemMap = new Hashtable();
            ItemMapCic = new Hashtable();
        }

        //get all workitems from tfs
        private WorkItemCollection GetWorkItemCollection()
        {
            WorkItemCollection workItemCollection = Store.Query(" SELECT * " +
                                                                  " FROM WorkItems " +
                                                                  " WHERE [System.TeamProject] = '" + _projectName +
                                                                  "' ORDER BY [System.Id]");
            return workItemCollection;
        }


        public void updateToLatestStatus(WorkItem oldWorkItem, WorkItem newWorkItem)
        {
            Queue<string> result = new Queue<string>();
            string previousState = null;
            string originalState = (string)newWorkItem.Fields["State"].Value;
            string sourceState = (string)oldWorkItem.Fields["State"].Value;
            string sourceFinalReason = (string)oldWorkItem.Fields["Reason"].Value;

            //try to change the status directly
            newWorkItem.Open();
            newWorkItem.Fields["State"].Value = oldWorkItem.Fields["State"].Value;
            //System.Diagnostics.Debug.WriteLine(newWorkItem.Type.Name + "      " + newWorkItem.Fields["State"].Value);

            //if status can't be changed directly... 
            if (newWorkItem.Fields["State"].Status != FieldStatus.Valid)
            {
                //get the state transition history of the source work item.
                foreach (Revision revision in oldWorkItem.Revisions)
                {
                    // Get Status          
                    if (!revision.Fields["State"].Value.Equals(previousState))
                    {
                        previousState = revision.Fields["State"].Value.ToString();
                        result.Enqueue(previousState);
                    }
                }

                int i = 1;
                previousState = originalState;
                //traverse new work item through old work items's transition states
                foreach (String currentStatus in result)
                {
                    bool success;
                    if (i != result.Count)
                    {
                        success = ChangeWorkItemStatus(newWorkItem, previousState, currentStatus);
                        previousState = currentStatus;
                    }
                    else
                    {
                        success = ChangeWorkItemStatus(newWorkItem, previousState, currentStatus, sourceFinalReason);
                    }
                    i++;
                    // If we could not do the incremental state change then we are done.  We will have to go back to the orginal...
                    if (!success)
                        break;
                }
            }
            else
            {
                // Just save it off if we can.
                ChangeWorkItemStatus(newWorkItem, originalState, sourceState);
            }
        }

        private bool ChangeWorkItemStatus(WorkItem workItem, string orginalSourceState, string destState)
        {
            //Try to save the new state.  If that fails then we also go back to the orginal state.
            try
            {
                workItem.Open();
                workItem.Fields["State"].Value = destState;
                workItem.Save();
                return true;
            }
            catch (Exception)
            {
                Logger.WarnFormat("Failed to save state for workItem: {0}  type:'{1}' state from '{2}' to '{3}' => rolling workItem status to original state '{4}'",
                    workItem.Id, workItem.Type.Name, orginalSourceState, destState, orginalSourceState);
                //Revert back to the original value.
                workItem.Fields["State"].Value = orginalSourceState;
                return false;
            }
        }

        //save final state transition and set final reason.
        private bool ChangeWorkItemStatus(WorkItem workItem, string orginalSourceState, string destState, string reason)
        {
            //Try to save the new state.  If that fails then we also go back to the orginal state.
            try
            {
                workItem.Open();
                workItem.Fields["State"].Value = destState;
                workItem.Fields["Reason"].Value = reason;

                workItem.Validate();
                workItem.Save();

                return true;
            }
            catch (Exception)
            {
                Logger.WarnFormat("Failed to save state for workItem: {0}  type:'{1}' state from '{2}' to '{3}' =>rolling workItem status to original state '{4}'",
                    workItem.Id, workItem.Type.Name, orginalSourceState, destState, orginalSourceState);
                //Revert back to the original value.
                workItem.Fields["State"].Value = orginalSourceState;
                return false;
            }
        }


        /* Copy work items to project from work item collection */
        public void writeWorkItems(WorkItemStore sourceStore, WorkItemCollection workItemCollection, string sourceProjectName, ProgressBar progressBar, Hashtable fieldMapAll, bool areVersionHistoryCommentsIncluded, bool shouldWorkItemsBeLinkedToGitCommits)
        {
            ReadItemMap(sourceProjectName);
            int i = 1;
            List<WorkItem> newItems = new List<WorkItem>();
            foreach (WorkItem workItem in workItemCollection)
            {
                if (ItemMap.ContainsKey(workItem.Id))
                {
                    continue;
                }

                WorkItem newWorkItem;
                Hashtable fieldMap = ListToTable((List<object>)fieldMapAll[workItem.Type.Name]);
                if (_workItemTypes.Contains(workItem.Type.Name))
                {
                    newWorkItem = new WorkItem(_workItemTypes[workItem.Type.Name]);
                }
                else if (workItem.Type.Name == "User Story")
                {
                    newWorkItem = new WorkItem(_workItemTypes["Product Backlog Item"]);
                }
                else if (workItem.Type.Name == "Issue")
                {
                    newWorkItem = new WorkItem(_workItemTypes["Impediment"]);
                }
                else
                {
                    Logger.InfoFormat("Work Item Type {0} does not exist in target TFS", workItem.Type.Name);
                    continue;
                }

                /* assign relevent fields*/
                foreach (Field field in workItem.Fields)
                {
                    if (field.Name.Contains("ID") || field.Name.Contains("State") || field.Name.Contains("Reason") || field.Name == "History")
                    {
                        continue;
                    }

                    if (newWorkItem.Fields.Contains(field.Name) && newWorkItem.Fields[field.Name].IsEditable)
                    {
                        newWorkItem.Fields[field.Name].Value = field.Value;
                        if (field.Name == "Iteration Path" || field.Name == "Area Path" || field.Name == "Node Name" || field.Name == "Team Project")
                        {
                            try
                            {
                                string itPath = (string)field.Value;
                                int length = sourceProjectName.Length;
                                string itPathNew = _destinationProject.Name + itPath.Substring(length);
                                newWorkItem.Fields[field.Name].Value = itPathNew;
                            }
                            catch
                            {
                                // ignored
                            }
                        }
                    }
                    //Add values to mapped fields
                    else if (fieldMap.ContainsKey(field.Name))
                    {
                        newWorkItem.Fields[(string)fieldMap[field.Name]].Value = field.Value;
                    }
                }

                // TODO: Commits
                /* Validate Item Before Save*/
                ArrayList array = newWorkItem.Validate();
                foreach (Field item in array)
                {
                    Logger.Info(String.Format("Work item {0} Validation Error in field: {1}  : {2}", workItem.Id, item.Name, newWorkItem.Fields[item.Name].Value));
                }
                //if work item is valid
                if (array.Count == 0)
                {
                    UploadAttachments(newWorkItem, workItem);

                    newWorkItem.History = "Original Work Item ID: " + workItem.Id;
                    newWorkItem.Save();

                    if (areVersionHistoryCommentsIncluded)
                    {
                        migrateComments(workItem, newWorkItem);
                    }


                    ItemMap.Add(workItem.Id, newWorkItem.Id);
                    newItems.Add(workItem);
                    //update workitem status
                    updateToLatestStatus(workItem, newWorkItem);
                }
                else
                {
                    Logger.ErrorFormat("Work item {0} could not be saved", workItem.Id);
                }

                var i1 = i;
                progressBar.Dispatcher.BeginInvoke(new Action(delegate
                {
                    progressBar.Value = i1 / (float)workItemCollection.Count * 100;
                }));
                i++;
            }

            WriteMaptoFile(sourceProjectName);
            CreateLinks(newItems, sourceStore, progressBar);

            CreateExternalLinks(newItems, sourceStore, progressBar);
        }

        private static void migrateComments(WorkItem workItem, WorkItem newWorkItem)
        {
            const int maxHistoryLength = 1048576;

            foreach (Revision revision in workItem.Revisions)
            {
                if (ShouldRevisionBeMigrated(revision))
                {
                    var history =
                        $"{revision.Fields["History"].Value}<br><p><FONT color=#808080 size=1><EM>Changed date: {revision.Fields["Changed Date"].Value}</EM></FONT></p>";
                    RemoveImagesIfHistoryIsTooLong(ref history);

                    if (history.Length < maxHistoryLength)
                    {
                        newWorkItem.Fields["Changed By"].Value = revision.Fields["Changed By"].Value;
                        newWorkItem.History = history;
                        newWorkItem.Save();
                    }
                }
            }

            bool ShouldRevisionBeMigrated(Revision revision)
            {
                return !string.IsNullOrWhiteSpace(revision.Fields["History"].Value?.ToString()) &&
                       revision.Fields["Changed By"].Value?.ToString() != "_SYSTFSBuild";
            }

            void RemoveImagesIfHistoryIsTooLong(ref string history)
            {
                int ImageStart(string s)
                {
                    return s.IndexOf("<img", StringComparison.Ordinal);
                }

                int ImageEnd(string s, int imageStart)
                {
                    return s.IndexOf(">", imageStart, StringComparison.Ordinal);
                }

                {
                    var imageStart = ImageStart(history);

                    while (history.Length >= maxHistoryLength && imageStart >= 0)
                    {
                        var imageEnd = ImageEnd(history, imageStart);
                        if (imageEnd < 0)
                            return;
                        history = history.Remove(imageStart, imageEnd - imageStart);
                        history = history.Insert(imageStart, "<p><FONT color=#808080 size=1><EM>(image removed)</EM></FONT></p>");
                        imageStart = ImageStart(history);
                    }
                }
            }


        }


        private Hashtable ListToTable(List<object> map)
        {
            Hashtable table = new Hashtable();
            if (map != null)
            {
                foreach (object[] item in map)
                {
                    try
                    {
                        table.Add((string)item[0], (string)item[1]);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("Error in ListToTable", ex);
                    }
                }
            }
            return table;
        }

        private void ReadItemMap(string sourceProjectName)
        {
            string filaPath = String.Format(@"Map\ID_map_{0}_to_{1}.txt", sourceProjectName, _projectName);
            ItemMap = new Hashtable();
            string line;
            if (File.Exists(filaPath))
            {
                StreamReader file = new StreamReader(filaPath);
                while ((line = file.ReadLine()) != null)
                {
                    if (line.Contains("Source ID|Target ID"))
                    {
                        continue;
                    }
                    string[] idMap = line.Split('|');
                    if (idMap[0].Trim() != "" && idMap[1].Trim() != "")
                    {
                        ItemMap.Add(Convert.ToInt32(idMap[0].Trim()), Convert.ToInt32(idMap[1].Trim()));
                    }
                }
                file.Close();
            }
        }

        /* Set links between workitems */
        private void CreateLinks(List<WorkItem> workItemCollection, WorkItemStore sourceStore, ProgressBar progressBar)
        {
            List<int> linkedWorkItemList = new List<int>();
            int index = 0;
            foreach (WorkItem workItem in workItemCollection)
            {
                WorkItemLinkCollection links = workItem.WorkItemLinks;
                if (links.Count > 0)
                {
                    int newWorkItemID = (int)ItemMap[workItem.Id];
                    WorkItem newWorkItem = Store.GetWorkItem(newWorkItemID);

                    foreach (WorkItemLink link in links)
                    {
                        try
                        {
                            WorkItem targetItem = sourceStore.GetWorkItem(link.TargetId);
                            if (ItemMap.ContainsKey(link.TargetId) && targetItem != null)
                            {
                                int targetWorkItemID = 0;
                                if (ItemMap.ContainsKey(link.TargetId))
                                {
                                    targetWorkItemID = (int)ItemMap[link.TargetId];
                                }

                                //if the link is not already created(check if target id is not in list)
                                if (!linkedWorkItemList.Contains(link.TargetId))
                                {
                                    try
                                    {
                                        WorkItemLinkTypeEnd linkTypeEnd = Store.WorkItemLinkTypes.LinkTypeEnds[link.LinkTypeEnd.Name];
                                        newWorkItem.Links.Add(new RelatedLink(linkTypeEnd, targetWorkItemID));

                                        ArrayList array = newWorkItem.Validate();
                                        if (array.Count == 0)
                                        {
                                            newWorkItem.Save();
                                        }
                                        else
                                        {
                                            Logger.Info("WorkItem Validation failed at link setup for work item: " + workItem.Id);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.ErrorFormat("Error occured when crearting link for work item: {0} target item: {1}", workItem.Id, link.TargetId);
                                        Logger.Error("Error detail", ex);
                                    }
                                }
                            }
                            else
                            {
                                Logger.Info("Link is not created for work item: " + workItem.Id + " - target item: " + link.TargetId + " does not exist");
                            }
                        }
                        catch (Exception)
                        {
                            Logger.Warn("Link is not created for work item: " + workItem.Id + " - target item: " + link.TargetId + " is not in Source TFS or you do not have permission to access");
                        }
                    }
                    //add the work item to list if the links are processed
                    linkedWorkItemList.Add(workItem.Id);
                }
                index++;
                var index1 = index;
                progressBar.Dispatcher.BeginInvoke(new Action(delegate
                {
                    progressBar.Value = index1 / (float)workItemCollection.Count * 100;
                }));
            }
        }

        private void CreateExternalLinks(List<WorkItem> workItemCollection, WorkItemStore sourceStore, ProgressBar progressBar)
        {
            int index = 0;
            foreach (WorkItem workItem in workItemCollection)
            {
                LinkCollection links = workItem.Links;
                if (links.Count > 0)
                {
                    int newWorkItemID = (int)ItemMap[workItem.Id];
                    WorkItem newWorkItem = Store.GetWorkItem(newWorkItemID);

                    var oldProjectName = string.Format("{0}%2F{1}", sourceStore.TeamProjectCollection.Name, workItem.Project.Name).Replace("tfs\\", "");
                    var newProjectName = string.Format("{0}%2F{1}", Store.TeamProjectCollection.Name, newWorkItem.Project.Name).Replace("tfs\\", "");
                    foreach (Link link in links)
                    {
                        try
                        {
                            var linkType = Store.RegisteredLinkTypes[link.ArtifactLinkType.Name];

                            if (link is ExternalLink)
                            {
                                //DON'T COPY CHANGESET LINKS
                                var oldLink = link as ExternalLink;
                                var uri = oldLink.LinkedArtifactUri;
                                if (!uri.ToLower().Contains("changeset"))
                                {
                                    uri = uri.Replace(oldProjectName, newProjectName);
                                    var newLink = new ExternalLink(linkType, uri);
                                    newWorkItem.Links.Add(newLink);
                                }
                            }
                            else if (link is Hyperlink)
                            {
                                var oldLink = link as Hyperlink;
                                var uri = oldLink.Location;
                                uri = uri.Replace(oldProjectName, newProjectName);
                                var newLink = new Hyperlink(uri);
                                newWorkItem.Links.Add(newLink);
                            }
                        }
                        catch (Exception)
                        {
                            Logger.Warn("Link is not created for work item: " + workItem.Id + " - target item: " + link.Comment + " is not in Source TFS or you do not have permission to access");
                        }
                    }
                    if (newWorkItem.IsDirty)
                    {
                        newWorkItem.Save();
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


        /* Upload attachments to workitems from local folder */
        private void UploadAttachments(WorkItem workItem, WorkItem workItemOld)
        {
            AttachmentCollection attachmentCollection = workItemOld.Attachments;
            foreach (Attachment att in attachmentCollection)
            {
                string comment = att.Comment;
                string name = @"Attachments\" + workItemOld.Id + "\\" + att.Name;
                string nameWithID = @"Attachments\" + workItemOld.Id + "\\" + att.Id + "_" + att.Name;
                try
                {
                    if (File.Exists(nameWithID))
                    {
                        workItem.Attachments.Add(new Attachment(nameWithID, comment));
                    }
                    else
                    {
                        workItem.Attachments.Add(new Attachment(name, comment));
                    }
                }
                catch (Exception ex)
                {
                    Logger.ErrorFormat("Error saving attachment: {0} for workitem: {1}", att.Name, workItemOld.Id);
                    Logger.Error("Error detail: ", ex);
                }
            }
        }


        public void GenerateIterations(XmlNode tree, string sourceProjectName)
        {
            ICommonStructureService4 css = (ICommonStructureService4)_tfs.GetService(typeof(ICommonStructureService4));
            string rootNodePath = string.Format("\\{0}\\Iteration", _projectName);
            var pathRoot = css.GetNodeFromPath(rootNodePath);

            if (tree.FirstChild != null)
            {
                var firstChild = tree.FirstChild;
                CreateIterationNodes(firstChild, css, pathRoot);
            }
            RefreshCache();
        }

        private static void CreateIterationNodes(XmlNode node, ICommonStructureService4 css,
            NodeInfo pathRoot)
        {
            int myNodeCount = node.ChildNodes.Count;
            for (int i = 0; i < myNodeCount; i++)
            {
                XmlNode childNode = node.ChildNodes[i];
                NodeInfo createdNode;
                var name = childNode.Attributes?["Name"].Value;
                try
                {
                    var uri = css.CreateNode(name, pathRoot.Uri);
                    Console.WriteLine("NodeCreated:" + uri);
                    createdNode = css.GetNode(uri);
                }
                catch (Exception)
                {
                    //node already exists
                    createdNode = css.GetNodeFromPath(pathRoot.Path + @"\" + name);
                    //continue;
                }
                DateTime? startDateToUpdate = null;
                if (!createdNode.StartDate.HasValue)
                {
                    var startDate = childNode.Attributes?["StartDate"];
                    DateTime startDateParsed;
                    if (startDate != null && DateTime.TryParse(startDate.Value, out startDateParsed))
                        startDateToUpdate = startDateParsed;
                }
                DateTime? finishDateToUpdate = null;
                if (!createdNode.FinishDate.HasValue)
                {
                    DateTime finishDateParsed;
                    var finishDate = childNode.Attributes?["FinishDate"];
                    if (finishDate != null && DateTime.TryParse(finishDate.Value, out finishDateParsed))
                        finishDateToUpdate = finishDateParsed;
                }
                if (startDateToUpdate.HasValue || finishDateToUpdate.HasValue)
                    css.SetIterationDates(createdNode.Uri, startDateToUpdate, finishDateToUpdate);
                if (node.HasChildNodes)
                {
                    foreach (XmlNode subChildNode in childNode.ChildNodes)
                    {
                        CreateIterationNodes(subChildNode, css, createdNode);
                    }
                }
            }
        }

        public void GenerateAreas(XmlNode tree, string sourceProjectName)
        {
            ICommonStructureService css = (ICommonStructureService)_tfs.GetService(typeof(ICommonStructureService));
            // get project info
            ProjectInfo projectInfo = css.GetProjectFromName(_projectName);
            NodeInfo[] nodes = css.ListStructures(projectInfo.Uri);

            // find ProjectModelHierarchy (contains path for area node)
            var node = nodes.FirstOrDefault(n => n.StructureType == "ProjectModelHierarchy");

            if (node == null)
                return;

            var pathRoot = css.GetNodeFromPath(node.Path);

            if (tree.FirstChild != null)
            {
                int myNodeCount = tree.FirstChild.ChildNodes.Count;
                for (int i = 0; i < myNodeCount; i++)
                {
                    XmlNode childNode = tree.ChildNodes[0].ChildNodes[i];
                    try
                    {
                        css.CreateNode(childNode.Attributes?["Name"].Value, pathRoot.Uri);
                    }
                    catch (Exception)
                    {
                        //node already exists
                        continue;
                    }
                    if (childNode.FirstChild != null)
                    {
                        string nodePath = node.Path + "\\" + childNode.Attributes?["Name"].Value;
                        GenerateSubAreas(childNode, nodePath, css);
                    }
                }
            }
            RefreshCache();
        }

        private void GenerateSubAreas(XmlNode tree, string nodePath, ICommonStructureService css)
        {
            var path = css.GetNodeFromPath(nodePath);
            int nodeCount = tree.FirstChild.ChildNodes.Count;
            for (int i = 0; i < nodeCount; i++)
            {
                XmlNode node = tree.ChildNodes[0].ChildNodes[i];
                try
                {
                    css.CreateNode(node.Attributes?["Name"].Value, path.Uri);
                }
                catch (Exception)
                {
                    //node already exists
                    continue;
                }
                if (node.FirstChild != null)
                {
                    string newPath = nodePath + "\\" + node.Attributes?["Name"].Value;
                    GenerateSubAreas(node, newPath, css);
                }
            }
        }

        private void RefreshCache()
        {
            ICommonStructureService css = _tfs.GetService<ICommonStructureService>();
            WorkItemServer server = _tfs.GetService<WorkItemServer>();
            server.SyncExternalStructures(WorkItemServer.NewRequestId(), css.GetProjectFromName(_projectName).Uri);
            Store.RefreshCache();
        }


        /* write ID mapping to local file */
        public void WriteMaptoFile(string sourceProjectName)
        {
            string filaPath = String.Format(@"Map\ID_map_{0}_to_{1}.txt", sourceProjectName, _projectName);
            if (!Directory.Exists(@"Map"))
            {
                Directory.CreateDirectory(@"Map");
            }
            else if (File.Exists(filaPath))
            {
                File.WriteAllText(filaPath, string.Empty);
            }

            using (StreamWriter file = new StreamWriter(filaPath, false))
            {
                file.WriteLine("Source ID|Target ID");
                foreach (object key in ItemMap)
                {
                    DictionaryEntry item = (DictionaryEntry)key;
                    file.WriteLine(item.Key + "\t | \t" + item.Value);
                }
            }
        }


        //Delete all workitems in project
        public void DeleteWorkItems()
        {
            WorkItemCollection workItemCollection = GetWorkItemCollection();
            List<int> toDeletes = new List<int>();

            foreach (WorkItem workItem in workItemCollection)
            {
                Debug.WriteLine(workItem.Id);
                toDeletes.Add(workItem.Id);
            }
            var errors = Store.DestroyWorkItems(toDeletes);
            foreach (var error in errors)
            {
                Debug.WriteLine(error.Exception.Message);
            }
        }

        /* Compare work item type definitions and add fields from source work item types and replace workflow */
        public void SetFieldDefinitions(WorkItemTypeCollection workItemTypesSource, Hashtable fieldList)
        {
            foreach (WorkItemType workItemTypeSource in workItemTypesSource)
            {
                WorkItemType workItemTypeTarget;
                if (workItemTypeSource.Name == "User Story")
                {
                    workItemTypeTarget = _workItemTypes["Product Backlog Item"];
                }
                else if (workItemTypeSource.Name == "Issue")
                {
                    workItemTypeTarget = _workItemTypes["Impediment"];
                }
                else
                {
                    workItemTypeTarget = _workItemTypes[workItemTypeSource.Name];
                }

                XmlDocument workItemTypeXmlSource = workItemTypeSource.Export(false);
                XmlDocument workItemTypeXmlTarget = workItemTypeTarget.Export(false);

                workItemTypeXmlTarget = AddNewFields(workItemTypeXmlSource, workItemTypeXmlTarget, (List<object>)fieldList[workItemTypeTarget.Name]);

                try
                {
                    WorkItemType.Validate(Store.Projects[_projectName], workItemTypeXmlTarget.InnerXml);
                    Store.Projects[_projectName].WorkItemTypes.Import(workItemTypeXmlTarget.InnerXml);
                }
                catch (XmlException)
                {
                    Logger.Info("XML import falied for " + workItemTypeSource.Name);
                }
            }
        }

        /* Add field definitions from Source xml to target xml */
        private XmlDocument AddNewFields(XmlDocument workItemTypeXmlSource, XmlDocument workItemTypeXmlTarget, List<object> fieldList)
        {
            XmlNodeList parentNodeList = workItemTypeXmlTarget.GetElementsByTagName("FIELDS");
            XmlNode parentNode = parentNodeList[0];
            foreach (object[] list in fieldList)
            {
                if ((bool)list[1])
                {
                    XmlNodeList transitionsListSource = workItemTypeXmlSource.SelectNodes("//FIELD[@name='" + list[0] + "']");
                    try
                    {
                        if (transitionsListSource != null)
                        {
                            XmlNode copiedNode = workItemTypeXmlTarget.ImportNode(transitionsListSource[0], true);
                            parentNode.AppendChild(copiedNode);
                        }
                    }
                    catch (Exception)
                    {
                        Logger.ErrorFormat("Error adding new field for parent node : {0}", parentNode.Value);
                    }
                }
            }
            return workItemTypeXmlTarget;
        }


        public string ReplaceWorkFlow(WorkItemTypeCollection workItemTypesSource, List<object> fieldList)
        {
            string error = "";
            for (int i = 0; i < fieldList.Count; i++)
            {
                object[] list = (object[])fieldList[i];
                if ((bool)list[1])
                {
                    WorkItemType workItemTypeTarget = _workItemTypes[(string)list[0]];

                    WorkItemType workItemTypeSource = null;
                    if (workItemTypesSource.Contains((string)list[0]))
                    {
                        workItemTypeSource = workItemTypesSource[(string)list[0]];
                    }
                    else if (workItemTypeTarget.Name == "Product Backlog Item")
                    {
                        workItemTypeSource = workItemTypesSource["User Story"];
                    }
                    else if (workItemTypeTarget.Name == "Impediment")
                    {
                        workItemTypeSource = workItemTypesSource["Issue"];
                    }

                    XmlDocument workItemTypeXmlSource = workItemTypeSource?.Export(false);
                    XmlDocument workItemTypeXmlTarget = workItemTypeTarget.Export(false);

                    if (workItemTypeXmlSource != null)
                    {
                        XmlNodeList transitionsListSource = workItemTypeXmlSource.GetElementsByTagName("WORKFLOW");
                        XmlNode transitions = transitionsListSource[0];

                        XmlNodeList transitionsListTarget = workItemTypeXmlTarget.GetElementsByTagName("WORKFLOW");
                        XmlNode transitionsTarget = transitionsListTarget[0];
                        string defTarget;
                        try
                        {
                            string def = workItemTypeXmlTarget.InnerXml;
                            string workflowSource = transitions.OuterXml;
                            string workflowTarget = transitionsTarget.OuterXml;

                            defTarget = def.Replace(workflowTarget, workflowSource);
                            WorkItemType.Validate(Store.Projects[_projectName], defTarget);
                            Store.Projects[_projectName].WorkItemTypes.Import(defTarget);
                            fieldList.Remove(list);
                            i--;
                        }
                        catch (Exception ex)
                        {
                            Logger.Error("Error Replacing work flow");
                            error = error + "Error Replacing work flow for " + (string)list[0] + ":" + ex.Message + "\n";
                        }
                    }
                }
            }
            return error;
        }


        private readonly string[] _sharedQueriesString = {
            "Shared Queries",
            "Freigegebene Abfragen"
        };
        public void SetTeamQueries(QueryHierarchy sourceQueryCol, string sourceProjectName)
        {
            foreach (var queryItem in sourceQueryCol)
            {
                var queryFolder = (QueryFolder) queryItem;
                if (queryFolder.Name == "Team Queries" || queryFolder.Name == "Shared Queries")
                {
                    QueryFolder teamQueriesFolder = null;

                    foreach (var localSharedQueryString in _sharedQueriesString)
                    {
                        if (Store.Projects[_projectName].QueryHierarchy.Contains(localSharedQueryString))
                            teamQueriesFolder = (QueryFolder)Store.Projects[_projectName].QueryHierarchy[localSharedQueryString];
                    }

                    if (teamQueriesFolder != null)
                        SetQueryItem(queryFolder, teamQueriesFolder, sourceProjectName);
                }
            }
        }

        private void SetQueryItem(QueryFolder queryFolder, QueryFolder parentFolder, string sourceProjectName)
        {
            QueryItem newItem = null;
            foreach (QueryItem subQuery in queryFolder)
            {
                try
                {
                    if (subQuery.GetType() == typeof(QueryFolder))
                    {
                        newItem = new QueryFolder(subQuery.Name);
                        if (!parentFolder.Contains(subQuery.Name))
                        {
                            parentFolder.Add(newItem);
                            Store.Projects[_projectName].QueryHierarchy.Save();
                            SetQueryItem((QueryFolder)subQuery, (QueryFolder)newItem, sourceProjectName);
                        }
                        else
                        {
                            Logger.WarnFormat("Query Folder {0} already exists", subQuery);
                        }
                    }
                    else
                    {
                        QueryDefinition oldDef = (QueryDefinition)subQuery;
                        string queryText = oldDef.QueryText.Replace(sourceProjectName, _projectName).Replace("User Story", "Product Backlog Item").Replace("Issue", "Impediment");

                        newItem = new QueryDefinition(subQuery.Name, queryText);
                        if (!parentFolder.Contains(subQuery.Name))
                        {
                            parentFolder.Add(newItem);
                            Store.Projects[_projectName].QueryHierarchy.Save();
                        }
                        else
                        {
                            Logger.WarnFormat("Query Definition {0} already exists", subQuery);
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (newItem != null)
                        newItem.Delete();
                    Logger.ErrorFormat("Error creating Query: {0} : {1}", subQuery, ex.Message);
                }
            }
        }

        public Hashtable MapFields(WorkItemTypeCollection workItemTypesSource)
        {
            Hashtable fieldMap = new Hashtable();

            foreach (WorkItemType workItemTypeSource in workItemTypesSource)
            {
                List<List<string>> fieldList = new List<List<string>>();
                List<string> sourceList = new List<string>();
                List<string> targetList = new List<string>();

                WorkItemType workItemTypeTarget;
                if (_workItemTypes.Contains(workItemTypeSource.Name))
                {
                    workItemTypeTarget = _workItemTypes[workItemTypeSource.Name];
                }
                else if (workItemTypeSource.Name == "User Story")
                {
                    workItemTypeTarget = _workItemTypes["Product Backlog Item"];
                }
                else if (workItemTypeSource.Name == "Issue")
                {
                    workItemTypeTarget = _workItemTypes["Impediment"];
                }
                else
                {
                    continue;
                }

                XmlDocument workItemTypeXmlSource = workItemTypeSource.Export(false);
                XmlDocument workItemTypeXmlTarget = workItemTypeTarget.Export(false);

                XmlNodeList fieldListSource = workItemTypeXmlSource.GetElementsByTagName("FIELD");
                XmlNodeList fieldListTarget = workItemTypeXmlTarget.GetElementsByTagName("FIELD");

                foreach (XmlNode field in fieldListSource)
                {
                    if (field.Attributes?["name"] != null)
                    {
                        XmlNodeList tempList = workItemTypeXmlTarget.SelectNodes("//FIELD[@name='" + field.Attributes["name"].Value + "']");
                        if (tempList != null && tempList.Count == 0)
                        {
                            sourceList.Add(field.Attributes["name"].Value);
                        }
                    }
                }
                fieldList.Add(sourceList);

                foreach (XmlNode field in fieldListTarget)
                {
                    if (field.Attributes?["name"] != null)
                    {
                        XmlNodeList tempList = workItemTypeXmlSource.SelectNodes("//FIELD[@name='" + field.Attributes["name"].Value + "']");
                        if (tempList != null && tempList.Count == 0)
                        {
                            targetList.Add(field.Attributes["name"].Value);
                        }
                    }
                }
                fieldList.Add(targetList);
                fieldMap.Add(workItemTypeTarget.Name, fieldList);
            }
            return fieldMap;
        }
    }
}
