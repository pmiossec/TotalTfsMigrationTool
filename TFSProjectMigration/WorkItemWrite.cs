﻿using System;
using System.Linq;
using System.Collections.Generic;
using System.Collections;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Proxy;
using System.Xml;
using System.IO;
using log4net;
using System.Windows.Controls;

namespace TFSProjectMigration
{
    public class WorkItemWrite
    {
        TfsTeamProjectCollection tfs;
        public WorkItemStore store;
        public QueryHierarchy queryCol;
        Project destinationProject;
        String projectName;
        public XmlNode AreaNodes;
        public XmlNode IterationsNodes;
        WorkItemTypeCollection workItemTypes;
        public Hashtable itemMap;
        public Hashtable itemMapCIC;
        private static readonly ILog logger = LogManager.GetLogger(typeof(TFSWorkItemMigrationUI));

        public WorkItemWrite(TfsTeamProjectCollection tfs, Project destinationProject)
        {
            this.tfs = tfs;
            projectName = destinationProject.Name;
            this.destinationProject = destinationProject;
            //store = (WorkItemStore)tfs.GetService(typeof(WorkItemStore));
            store = new WorkItemStore(tfs, WorkItemStoreFlags.BypassRules);
            queryCol = store.Projects[destinationProject.Name].QueryHierarchy;
            workItemTypes = store.Projects[destinationProject.Name].WorkItemTypes;
            itemMap = new Hashtable();
            itemMapCIC = new Hashtable();
        }

        //get all workitems from tfs
        private WorkItemCollection GetWorkItemCollection()
        {
            WorkItemCollection workItemCollection = store.Query(" SELECT * " +
                                                                  " FROM WorkItems " +
                                                                  " WHERE [System.TeamProject] = '" + projectName +
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
                    bool success = false;
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
                bool success = ChangeWorkItemStatus(newWorkItem, originalState, sourceState);
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
                logger.WarnFormat("Failed to save state for workItem: {0}  type:'{1}' state from '{2}' to '{3}' => rolling workItem status to original state '{4}'",
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

                ArrayList list = workItem.Validate();
                workItem.Save();

                return true;
            }
            catch (Exception)
            {
                logger.WarnFormat("Failed to save state for workItem: {0}  type:'{1}' state from '{2}' to '{3}' =>rolling workItem status to original state '{4}'",
                    workItem.Id, workItem.Type.Name, orginalSourceState, destState, orginalSourceState);
                //Revert back to the original value.
                workItem.Fields["State"].Value = orginalSourceState;
                return false;
            }
        }


        /* Copy work items to project from work item collection */
        public void writeWorkItems(WorkItemStore sourceStore, WorkItemCollection workItemCollection, string sourceProjectName, ProgressBar ProgressBar, Hashtable fieldMapAll)
        {
            ReadItemMap(sourceProjectName);
            int i = 1;
            List<WorkItem> newItems = new List<WorkItem>();
            foreach (WorkItem workItem in workItemCollection)
            {
                if (itemMap.ContainsKey(workItem.Id))
                {
                    continue;
                }

                WorkItem newWorkItem = null;
                Hashtable fieldMap = ListToTable((List<object>)fieldMapAll[workItem.Type.Name]);
                if (workItemTypes.Contains(workItem.Type.Name))
                {
                    newWorkItem = new WorkItem(workItemTypes[workItem.Type.Name]);
                }
                else if (workItem.Type.Name == "User Story")
                {
                    newWorkItem = new WorkItem(workItemTypes["Product Backlog Item"]);
                }
                else if (workItem.Type.Name == "Issue")
                {
                    newWorkItem = new WorkItem(workItemTypes["Impediment"]);
                }
                else
                {
                    logger.InfoFormat("Work Item Type {0} does not exist in target TFS", workItem.Type.Name);
                    continue;
                }

                /* assign relevent fields*/
                foreach (Field field in workItem.Fields)
                {
                    if (field.Name.Contains("ID") || field.Name.Contains("State") || field.Name.Contains("Reason"))
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
                                string itPathNew = destinationProject.Name + itPath.Substring(length);
                                newWorkItem.Fields[field.Name].Value = itPathNew;
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }
                    //Add values to mapped fields
                    else if (fieldMap.ContainsKey(field.Name))
                    {
                        newWorkItem.Fields[(string)fieldMap[field.Name]].Value = field.Value;
                    }
                }

                /* Validate Item Before Save*/
                ArrayList array = newWorkItem.Validate();
                foreach (Field item in array)
                {
                    logger.Info(String.Format("Work item {0} Validation Error in field: {1}  : {2}", workItem.Id, item.Name, newWorkItem.Fields[item.Name].Value));
                }
                //if work item is valid
                if (array.Count == 0)
                {
                    UploadAttachments(newWorkItem, workItem);
                    newWorkItem.Save();

                    itemMap.Add(workItem.Id, newWorkItem.Id);
                    newItems.Add(workItem);
                    //update workitem status
                    updateToLatestStatus(workItem, newWorkItem);
                }
                else
                {
                    logger.ErrorFormat("Work item {0} could not be saved", workItem.Id);
                }

                ProgressBar.Dispatcher.BeginInvoke(new Action(delegate ()
                {
                    float progress = (float)i / (float)workItemCollection.Count;
                    ProgressBar.Value = ((float)i / (float)workItemCollection.Count) * 100;
                }));
                i++;
            }

            WriteMaptoFile(sourceProjectName);
            CreateLinks(newItems, sourceStore, ProgressBar);
            
            CreateExternalLinks(newItems, sourceStore, ProgressBar);
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
                        logger.Error("Error in ListToTable", ex);
                    }
                }
            }
            return table;
        }

        private void ReadItemMap(string sourceProjectName)
        {
            string filaPath = String.Format(@"Map\ID_map_{0}_to_{1}.txt", sourceProjectName, projectName);
            itemMap = new Hashtable();
            string line;
            if (File.Exists(filaPath))
            {
                System.IO.StreamReader file = new System.IO.StreamReader(filaPath);
                while ((line = file.ReadLine()) != null)
                {
                    if (line.Contains("Source ID|Target ID"))
                    {
                        continue;
                    }
                    string[] idMap = line.Split(new char[] { '|' });
                    if (idMap[0].Trim() != "" && idMap[1].Trim() != "")
                    {
                        itemMap.Add(Convert.ToInt32(idMap[0].Trim()), Convert.ToInt32(idMap[1].Trim()));
                    }
                }
                file.Close();
            }
        }

        /* Set links between workitems */
        private void CreateLinks(List<WorkItem> workItemCollection, WorkItemStore sourceStore, ProgressBar ProgressBar)
        {
            List<int> linkedWorkItemList = new List<int>();
            WorkItemCollection targetWorkItemCollection = GetWorkItemCollection();
            int index = 0;
            foreach (WorkItem workItem in workItemCollection)
            {
                WorkItemLinkCollection links = workItem.WorkItemLinks;
                if (links.Count > 0)
                {
                    int newWorkItemID = (int)itemMap[workItem.Id];
                    WorkItem newWorkItem = store.GetWorkItem(newWorkItemID);

                    foreach (WorkItemLink link in links)
                    {
                        try
                        {
                            WorkItem targetItem = sourceStore.GetWorkItem(link.TargetId);
                            if (itemMap.ContainsKey(link.TargetId) && targetItem != null)
                            {

                                int targetWorkItemID = 0;
                                if (itemMap.ContainsKey(link.TargetId))
                                {
                                    targetWorkItemID = (int)itemMap[link.TargetId];
                                }

                                //if the link is not already created(check if target id is not in list)
                                if (!linkedWorkItemList.Contains(link.TargetId))
                                {
                                    try
                                    {
                                        WorkItemLinkTypeEnd linkTypeEnd = store.WorkItemLinkTypes.LinkTypeEnds[link.LinkTypeEnd.Name];
                                        newWorkItem.Links.Add(new RelatedLink(linkTypeEnd, targetWorkItemID));

                                        ArrayList array = newWorkItem.Validate();
                                        if (array.Count == 0)
                                        {
                                            newWorkItem.Save();
                                        }
                                        else
                                        {
                                            logger.Info("WorkItem Validation failed at link setup for work item: " + workItem.Id);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        logger.ErrorFormat("Error occured when crearting link for work item: {0} target item: {1}", workItem.Id, link.TargetId);
                                        logger.Error("Error detail", ex);
                                    }

                                }
                            }
                            else
                            {
                                logger.Info("Link is not created for work item: " + workItem.Id + " - target item: " + link.TargetId + " does not exist");
                            }
                        }
                        catch (Exception)
                        {
                            logger.Warn("Link is not created for work item: " + workItem.Id + " - target item: " + link.TargetId + " is not in Source TFS or you do not have permission to access");
                        }
                    }
                    //add the work item to list if the links are processed
                    linkedWorkItemList.Add(workItem.Id);

                }
                index++;
                ProgressBar.Dispatcher.BeginInvoke(new Action(delegate ()
                {
                    float progress = (float)index / (float)workItemCollection.Count;
                    ProgressBar.Value = ((float)index / (float)workItemCollection.Count) * 100;
                }));

            }
        }

        private void CreateExternalLinks(List<WorkItem> workItemCollection, WorkItemStore sourceStore, ProgressBar ProgressBar)
        {
            List<int> linkedWorkItemList = new List<int>();
            WorkItemCollection targetWorkItemCollection = GetWorkItemCollection();

            int index = 0;
            foreach (WorkItem workItem in workItemCollection)
            {
                LinkCollection links = workItem.Links;
                if (links.Count > 0)
                {
                    int newWorkItemID = (int)itemMap[workItem.Id];
                    WorkItem newWorkItem = store.GetWorkItem(newWorkItemID);

                    //for(int i=0;i<newWorkItem.Links.Count;i++)
                    //{
                    //    var link = newWorkItem.Links[i];
                    //    if(link is ExternalLink)
                    //    {
                    //        newWorkItem.Links.Remove(link);
                    //        i--;
                    //    }
                    //}
                    var oldProjectName = string.Format("{0}%2F{1}", sourceStore.TeamProjectCollection.Name, workItem.Project.Name).Replace("tfs\\","");
                    var newProjectName = string.Format("{0}%2F{1}", store.TeamProjectCollection.Name, newWorkItem.Project.Name).Replace("tfs\\", "");
                    foreach (Link link in links)
                    {
                        try
                        {

                            var linkType = store.RegisteredLinkTypes[link.ArtifactLinkType.Name];

                            if (link is ExternalLink)
                            {
                                //DON'T COPY CHANGESET LINKS
                                var oldLink = link as ExternalLink;
                                var uri = oldLink.LinkedArtifactUri;
                                if(!uri.ToLower().Contains("changeset"))
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
                            //else if (link is RelatedLink)
                            //{
                            //    var oldLink = link as RelatedLink;
                            //    var oldItemId = oldLink.RelatedWorkItemId;
                            //    if (itemMap.ContainsKey(oldItemId))
                            //    {
                            //        var linkTypeEnd = store.WorkItemLinkTypes.LinkTypeEnds[oldLink.LinkTypeEnd.Name];
                            //        var newLink = new RelatedLink(linkTypeEnd, (int)itemMap[oldItemId]);
                            //        newWorkItem.Links.Add(newLink);
                            //    }
                            //}

                        }
                        catch (Exception)
                        {
                            logger.Warn("Link is not created for work item: " + workItem.Id + " - target item: " + link.Comment + " is not in Source TFS or you do not have permission to access");
                        }
                    }
                    if (newWorkItem.IsDirty)
                    {
                        newWorkItem.Save();
                    }
                }

                index++;
                ProgressBar.Dispatcher.BeginInvoke(new Action(delegate ()
                {
                    float progress = (float)index / (float)workItemCollection.Count;
                    ProgressBar.Value = ((float)index / (float)workItemCollection.Count) * 100;
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
                    logger.ErrorFormat("Error saving attachment: {0} for workitem: {1}", att.Name, workItemOld.Id);
                    logger.Error("Error detail: ", ex);
                }
            }
        }


        public void GenerateIterations(XmlNode tree, string sourceProjectName)
        {
            ICommonStructureService4 css = (ICommonStructureService4)tfs.GetService(typeof(ICommonStructureService4));
            string rootNodePath = string.Format("\\{0}\\Iteration", projectName);
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
                var name = childNode.Attributes["Name"].Value;
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
                    var startDate = childNode.Attributes["StartDate"];
                    DateTime startDateParsed;
                    if (startDate != null && DateTime.TryParse(startDate.Value, out startDateParsed))
                        startDateToUpdate = startDateParsed;
                }
                DateTime? finishDateToUpdate = null;
                if (!createdNode.FinishDate.HasValue)
                {
                    DateTime finishDateParsed;
                    var finishDate = childNode.Attributes["FinishDate"];
                    if (finishDate != null && DateTime.TryParse(finishDate.Value, out finishDateParsed))
                        finishDateToUpdate = finishDateParsed;
                }
                if (startDateToUpdate.HasValue || finishDateToUpdate.HasValue)
                    css.SetIterationDates(createdNode.Uri, startDateToUpdate, finishDateToUpdate);
                if (createdNode != null && node.HasChildNodes)
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
            ICommonStructureService css = (ICommonStructureService)tfs.GetService(typeof(ICommonStructureService));
            string rootNodePath = string.Format("\\{0}\\Area", projectName);
            var pathRoot = css.GetNodeFromPath(rootNodePath);

            if (tree.FirstChild != null)
            {
                int myNodeCount = tree.FirstChild.ChildNodes.Count;
                for (int i = 0; i < myNodeCount; i++)
                {
                    XmlNode Node = tree.ChildNodes[0].ChildNodes[i];
                    try
                    {
                        css.CreateNode(Node.Attributes["Name"].Value, pathRoot.Uri);
                    }
                    catch (Exception)
                    {
                        //node already exists
                        continue;
                    }
                    if (Node.FirstChild != null)
                    {
                        string nodePath = rootNodePath + "\\" + Node.Attributes["Name"].Value;
                        GenerateSubAreas(Node, nodePath, css);
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
                    css.CreateNode(node.Attributes["Name"].Value, path.Uri);
                }
                catch (Exception ex)
                {
                    //node already exists
                    continue;
                }
                if (node.FirstChild != null)
                {
                    string newPath = nodePath + "\\" + node.Attributes["Name"].Value;
                    GenerateSubAreas(node, newPath, css);
                }
            }
        }

        private void RefreshCache()
        {
            ICommonStructureService css = tfs.GetService<ICommonStructureService>();
            WorkItemServer server = tfs.GetService<WorkItemServer>();
            server.SyncExternalStructures(WorkItemServer.NewRequestId(), css.GetProjectFromName(projectName).Uri);
            store.RefreshCache();
        }


        /* write ID mapping to local file */
        public void WriteMaptoFile(string sourceProjectName)
        {
            string filaPath = String.Format(@"Map\ID_map_{0}_to_{1}.txt", sourceProjectName, projectName);
            if (!Directory.Exists(@"Map"))
            {
                Directory.CreateDirectory(@"Map");
            }
            else if (File.Exists(filaPath))
            {
                System.IO.File.WriteAllText(filaPath, string.Empty);
            }

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(filaPath, false))
            {
                file.WriteLine("Source ID|Target ID");
                foreach (object key in itemMap)
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
                System.Diagnostics.Debug.WriteLine(workItem.Id);
                toDeletes.Add(workItem.Id);
            }
            var errors = store.DestroyWorkItems(toDeletes);
            foreach (var error in errors)
            {
                System.Diagnostics.Debug.WriteLine(error.Exception.Message);
            }

        }

        /* Compare work item type definitions and add fields from source work item types and replace workflow */
        public void SetFieldDefinitions(WorkItemTypeCollection workItemTypesSource, Hashtable fieldList)
        {
            foreach (WorkItemType workItemTypeSource in workItemTypesSource)
            {
                WorkItemType workItemTypeTarget = null;
                if (workItemTypeSource.Name == "User Story")
                {
                    workItemTypeTarget = workItemTypes["Product Backlog Item"];
                }
                else if (workItemTypeSource.Name == "Issue")
                {
                    workItemTypeTarget = workItemTypes["Impediment"];
                }
                else
                {
                    workItemTypeTarget = workItemTypes[workItemTypeSource.Name];
                }

                XmlDocument workItemTypeXmlSource = workItemTypeSource.Export(false);
                XmlDocument workItemTypeXmlTarget = workItemTypeTarget.Export(false);

                workItemTypeXmlTarget = AddNewFields(workItemTypeXmlSource, workItemTypeXmlTarget, (List<object>)fieldList[workItemTypeTarget.Name]);

                try
                {
                    WorkItemType.Validate(store.Projects[projectName], workItemTypeXmlTarget.InnerXml);
                    store.Projects[projectName].WorkItemTypes.Import(workItemTypeXmlTarget.InnerXml);
                }
                catch (XmlException)
                {
                    logger.Info("XML import falied for " + workItemTypeSource.Name);
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
                        XmlNode copiedNode = workItemTypeXmlTarget.ImportNode(transitionsListSource[0], true);
                        parentNode.AppendChild(copiedNode);
                    }
                    catch (Exception)
                    {
                        logger.ErrorFormat("Error adding new field for parent node : {0}", parentNode.Value);
                    }
                }
            }
            return workItemTypeXmlTarget;
        }


        /*Add new Field definition to work item type */
        private XmlDocument AddField(XmlDocument workItemTypeXml, string fieldName, string fieldRefName, string fieldType, string fieldReportable)
        {
            XmlNodeList tempList = workItemTypeXml.SelectNodes("//FIELD[@name='" + fieldName + "']");
            if (tempList.Count == 0)
            {
                XmlNode parent = workItemTypeXml.GetElementsByTagName("FIELDS")[0];
                XmlElement node = workItemTypeXml.CreateElement("FIELD");
                node.SetAttribute("name", fieldName);
                node.SetAttribute("refname", fieldRefName);
                node.SetAttribute("type", fieldType);
                node.SetAttribute("reportable", fieldReportable);
                parent.AppendChild(node);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Field already exists...");
                logger.InfoFormat("Field {0} already exists", fieldName);
            }
            return workItemTypeXml;
        }


        public string ReplaceWorkFlow(WorkItemTypeCollection workItemTypesSource, List<object> fieldList)
        {
            string error = "";
            for (int i = 0; i < fieldList.Count; i++)
            {
                object[] list = (object[])fieldList[i];
                if ((bool)list[1])
                {
                    WorkItemType workItemTypeTarget = workItemTypes[(string)list[0]];

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

                    XmlDocument workItemTypeXmlSource = workItemTypeSource.Export(false);
                    XmlDocument workItemTypeXmlTarget = workItemTypeTarget.Export(false);

                    XmlNodeList transitionsListSource = workItemTypeXmlSource.GetElementsByTagName("WORKFLOW");
                    XmlNode transitions = transitionsListSource[0];

                    XmlNodeList transitionsListTarget = workItemTypeXmlTarget.GetElementsByTagName("WORKFLOW");
                    XmlNode transitionsTarget = transitionsListTarget[0];
                    string defTarget = "";
                    try
                    {
                        string def = workItemTypeXmlTarget.InnerXml;
                        string workflowSource = transitions.OuterXml;
                        string workflowTarget = transitionsTarget.OuterXml;

                        defTarget = def.Replace(workflowTarget, workflowSource);
                        WorkItemType.Validate(store.Projects[projectName], defTarget);
                        store.Projects[projectName].WorkItemTypes.Import(defTarget);
                        fieldList.Remove(list);
                        i--;
                    }
                    catch (Exception ex)
                    {
                        logger.Error("Error Replacing work flow");
                        error = error + "Error Replacing work flow for " + (string)list[0] + ":" + ex.Message + "\n";
                    }

                }
            }
            return error;
        }


        private object[] GetAllTransitionsForWorkItemType(XmlDocument workItemTypeXml)
        {
            XmlNodeList transitionsList = workItemTypeXml.GetElementsByTagName("TRANSITION");

            string[] start = new string[transitionsList.Count];
            string[] dest = new string[transitionsList.Count];
            string[][] values = new string[transitionsList.Count][];

            int j = 0;
            foreach (XmlNode transition in transitionsList)
            {
                start[j] = transition.Attributes["from"].Value;
                dest[j] = transition.Attributes["to"].Value;

                XmlNodeList reasons = transition.SelectNodes("REASONS/REASON");

                string[] reasonVal = new string[1 + reasons.Count];
                reasonVal[0] = transition.SelectSingleNode("REASONS/DEFAULTREASON").Attributes["value"].Value;

                int i = 1;
                if (reasons != null)
                {
                    foreach (XmlNode reason in reasons)
                    {
                        reasonVal[i] = reason.Attributes["value"].Value;
                        i++;
                    }
                }
                values[j] = reasonVal;
                j++;
            }

            return new object[] { start, dest, values };
        }


        public void SetTeamQueries(QueryHierarchy sourceQueryCol, string sourceProjectName)
        {
            foreach (QueryFolder queryFolder in sourceQueryCol)
            {
                if (queryFolder.Name == "Team Queries" || queryFolder.Name == "Shared Queries")
                {
                    QueryFolder teamQueriesFolder = (QueryFolder)store.Projects[projectName].QueryHierarchy["Shared Queries"];
                    SetQueryItem(queryFolder, teamQueriesFolder, sourceProjectName);

                    QueryFolder test = (QueryFolder)store.Projects[projectName].QueryHierarchy["Shared Queries"];
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
                            store.Projects[projectName].QueryHierarchy.Save();
                            SetQueryItem((QueryFolder)subQuery, (QueryFolder)newItem, sourceProjectName);
                        }
                        else
                        {
                            logger.WarnFormat("Query Folder {0} already exists", subQuery);
                        }

                    }
                    else
                    {
                        QueryDefinition oldDef = (QueryDefinition)subQuery;
                        string queryText = oldDef.QueryText.Replace(sourceProjectName, projectName).Replace("User Story", "Product Backlog Item").Replace("Issue", "Impediment");

                        newItem = new QueryDefinition(subQuery.Name, queryText);
                        if (!parentFolder.Contains(subQuery.Name))
                        {
                            parentFolder.Add(newItem);
                            store.Projects[projectName].QueryHierarchy.Save();
                        }
                        else
                        {
                            logger.WarnFormat("Query Definition {0} already exists", subQuery);
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (newItem != null)
                        newItem.Delete();
                    logger.ErrorFormat("Error creating Query: {0} : {1}", subQuery, ex.Message);
                    continue;
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

                WorkItemType workItemTypeTarget = null;
                if (workItemTypes.Contains(workItemTypeSource.Name))
                {
                    workItemTypeTarget = workItemTypes[workItemTypeSource.Name];
                }
                else if (workItemTypeSource.Name == "User Story")
                {
                    workItemTypeTarget = workItemTypes["Product Backlog Item"];
                }
                else if (workItemTypeSource.Name == "Issue")
                {
                    workItemTypeTarget = workItemTypes["Impediment"];
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
                    if (field.Attributes["name"] != null)
                    {
                        XmlNodeList tempList = workItemTypeXmlTarget.SelectNodes("//FIELD[@name='" + field.Attributes["name"].Value + "']");
                        if (tempList.Count == 0)
                        {
                            sourceList.Add(field.Attributes["name"].Value);
                        }
                    }
                }
                fieldList.Add(sourceList);

                foreach (XmlNode field in fieldListTarget)
                {
                    if (field.Attributes["name"] != null)
                    {
                        XmlNodeList tempList = workItemTypeXmlSource.SelectNodes("//FIELD[@name='" + field.Attributes["name"].Value + "']");
                        if (tempList.Count == 0)
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
