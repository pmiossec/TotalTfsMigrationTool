
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.Server;
using System.ComponentModel;
using System.Xml;
using log4net.Config;
using log4net;
using System.Threading;

namespace TFSProjectMigration
{
    /// <summary>
    /// Interaction logic for TFSProjectMigrationUI.xaml
    /// </summary>
    public partial class TFSWorkItemMigrationUI : Window
    {

        private TfsTeamProjectCollection sourceTFS;
        private TfsTeamProjectCollection destinationTFS;
        private WorkItemStore sourceStore;
        private WorkItemStore destinationStore;
        private Project sourceProject;
        private Project destinationProject;
        private static readonly ILog logger = LogManager.GetLogger(typeof(TFSWorkItemMigrationUI));
        public int migrationState = 0;
        private bool IsNotIncludeClosed = false;
        private bool IsNotIncludeRemoved = false;
        WorkItemRead readSource;
        WorkItemWrite writeTarget;
        Hashtable fieldMap;
        Hashtable finalFieldMap;
        Hashtable copyingFieldSet;
        List<object> migrateTypeSet;

        public TFSWorkItemMigrationUI()
        {
            InitializeComponent();
            DOMConfigurator.Configure();
        }

        private void ConnectSourceProjectButton_Click(object sender, RoutedEventArgs e)
        {
            TeamProjectPicker tpp = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false);
            DialogResult result = tpp.ShowDialog();
            if (result.ToString() == "OK")
            {
                StatusViwer.Content = "";
                MigratingLabel.Content = "";
                StatusBar.Value = 0;
                SourceFieldGrid.ItemsSource = new List<string>();
                TargetFieldGrid.ItemsSource = new List<string>();
                MappedListGrid.ItemsSource = new List<object>();
                FieldToCopyGrid.ItemsSource = new List<object>();
                WorkFlowListGrid.ItemsSource = new List<object>();

                finalFieldMap = new Hashtable();
                copyingFieldSet = new Hashtable();
                migrateTypeSet = new List<object>();
                this.sourceTFS = tpp.SelectedTeamProjectCollection;
                this.sourceStore = (WorkItemStore)sourceTFS.GetService(typeof(WorkItemStore));

                this.sourceProject = sourceStore.Projects[tpp.SelectedProjects[0].Name];
                SourceProjectText.Text = string.Format("{0}/{1}", sourceTFS.Uri.ToString(), sourceProject.Name);
                readSource = new WorkItemRead(sourceTFS, sourceProject);

                if ((string)ConnectionStatusLabel.Content == "Select a Source project")
                {
                    ConnectionStatusLabel.Content = "";
                }

            }
        }

        private void ConnectDestinationProjectButton_Click(object sender, RoutedEventArgs e)
        {
            TeamProjectPicker tpp = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false);
            DialogResult result = tpp.ShowDialog();

            if (result.ToString() == "OK")
            {
                StatusViwer.Content = "";
                MigratingLabel.Content = "";
                StatusBar.Value = 0;
                SourceFieldGrid.ItemsSource = new List<string>();
                TargetFieldGrid.ItemsSource = new List<string>();
                MappedListGrid.ItemsSource = new List<object>();
                FieldToCopyGrid.ItemsSource = new List<object>();
                WorkFlowListGrid.ItemsSource = new List<object>();
                finalFieldMap = new Hashtable();
                copyingFieldSet = new Hashtable();
                migrateTypeSet = new List<object>();

                this.destinationTFS = tpp.SelectedTeamProjectCollection;
                this.destinationStore = (WorkItemStore)destinationTFS.GetService(typeof(WorkItemStore));

                this.destinationProject = destinationStore.Projects[tpp.SelectedProjects[0].Name];
                DestinationProjectText.Text = string.Format("{0}/{1}", destinationTFS.Uri.ToString(), destinationProject.Name);
                writeTarget = new WorkItemWrite(destinationTFS, destinationProject);

                if ((string)ConnectionStatusLabel.Content == "Select a Target project")
                {
                    ConnectionStatusLabel.Content = "";
                }

            }
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (SourceProjectText.Text.Equals(""))
            {
                ConnectionStatusLabel.Content = "Select a Source project";
            }
            else if (DestinationProjectText.Text.Equals(""))
            {
                ConnectionStatusLabel.Content = "Select a Target project";
            }
            else
            {
                fieldMap = writeTarget.MapFields(readSource.workItemTypes);
                IsNotIncludeClosed = (bool)ClosedTextBox.IsChecked;
                IsNotIncludeRemoved = (bool)RemovedTextBox.IsChecked;
                FieldCopyTab.IsEnabled = true;
                FieldCopyTab.IsSelected = true;
                
            }
        }


        public void projectMigration()
        {
            logger.Info(String.Format("--------------------------------Migration from {0} to {1} Start----------------------------------------------", sourceProject.Name, destinationProject.Name));
            CheckTestPlanTextBlock.Dispatcher.BeginInvoke(new Action(delegate()
            {
                CheckTestPlanTextBlock.Visibility = Visibility.Hidden;
            }));
            CheckLogTextBlock.Dispatcher.BeginInvoke(new Action(delegate()
            {
                CheckLogTextBlock.Visibility = Visibility.Hidden;
            }));
            MigratingLabel.Dispatcher.BeginInvoke(new Action(delegate()
            {
                MigratingLabel.Content = "Migrating...";
            }));

            StatusBar.Dispatcher.BeginInvoke(new Action(delegate()
            {
                StatusBar.Visibility = Visibility.Visible;
            }));
            WorkItemCollection source = readSource.GetWorkItems(sourceProject.Name, IsNotIncludeClosed, IsNotIncludeRemoved); //Get Workitems from source tfs 
            XmlNode[] iterations = readSource.PopulateIterations(); //Get Iterations and Areas from source tfs 

            StatusViwer.Dispatcher.BeginInvoke(new Action(delegate()
            {
                StatusViwer.Content = "Generating Areas...";
            }));
            writeTarget.GenerateAreas(iterations[0], sourceProject.Name); //Copy Areas

            StatusViwer.Dispatcher.BeginInvoke(new Action(delegate()
            {
                StatusViwer.Content = StatusViwer.Content + "\nGenerating Iterations...";
            }));
            writeTarget.GenerateIterations(iterations[1], sourceProject.Name); //Copy Iterations

            StatusViwer.Dispatcher.BeginInvoke(new Action(delegate()
            {
                StatusViwer.Content = StatusViwer.Content + "\nCopying Team Queries...";
            }));
            writeTarget.SetTeamQueries(readSource.queryCol, sourceProject.Name); //Copy Queries

            StatusViwer.Dispatcher.BeginInvoke(new Action(delegate()
            {
                StatusViwer.Content = StatusViwer.Content + "\nCopying Work Items...";
            }));
            writeTarget.writeWorkItems(sourceStore, source, sourceProject.Name, StatusBar, finalFieldMap); //Copy Workitems

            StatusViwer.Dispatcher.BeginInvoke(new Action(delegate()
            {
                StatusViwer.Content = StatusViwer.Content + "\nCopying Test Plans...";
            }));
            TestPlanMigration tcm = new TestPlanMigration(sourceTFS, destinationTFS, sourceProject.Name, destinationProject.Name, writeTarget.itemMap, StatusBar);
            tcm.CopyTestPlans(); //Copy Test Plans

            MigratingLabel.Dispatcher.BeginInvoke(new Action(delegate()
            {
                MigratingLabel.Content = "Project Migrated";
            }));

            StatusBar.Dispatcher.BeginInvoke(new Action(delegate()
            {
                StatusBar.Visibility = Visibility.Hidden;
            }));
            CheckTestPlanTextBlock.Dispatcher.BeginInvoke(new Action(delegate()
            {
                CheckTestPlanTextBlock.Visibility = Visibility.Visible;
            }));
            CheckLogTextBlock.Dispatcher.BeginInvoke(new Action(delegate()
            {
                CheckLogTextBlock.Visibility = Visibility.Visible;
            }));
            logger.Info("--------------------------------Migration END----------------------------------------------");
        }

        private void MigrationButton_Click(object sender, RoutedEventArgs e)
        {
            StartTab.IsEnabled = true;
            StartTab.IsSelected = true;
            StatusViwer.Content = "";
            MigratingLabel.Content = "";
            StatusBar.Value = 0;
            Thread migrationThread = new Thread(new ThreadStart(projectMigration));
            migrationThread.Start();
        }

        private void MigrationTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FieldMappingTab.IsSelected && (e.OriginalSource is System.Windows.Controls.TabControl))
            {
                SetFieldTypeList(fieldMap.Keys);
            }
            else if (FieldCopyTab.IsSelected && (e.OriginalSource is System.Windows.Controls.TabControl))
            {
                SetListsCopyFieldsTab(fieldMap.Keys);
                WorkFlowListGrid.ItemsSource = migrateTypeSet;
                WorkFlowListGrid.Items.Refresh();
            }
        }

        private void SelectedValueChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FieldTypesComboBox.SelectedValue != null)
            {
                SetFieldLists((string)FieldTypesComboBox.SelectedValue);
            }
        }

        private void SetFieldLists(string fieldType)
        {
            List<List<string>> fieldList = (List<List<string>>)fieldMap[fieldType];
            List<string> sourceList = (List<string>)fieldList.ElementAt(0);
            List<string> targetList = (List<string>)fieldList.ElementAt(1);

            SourceFieldGrid.ItemsSource = sourceList;
            TargetFieldGrid.ItemsSource = targetList;

            SourceFieldComboBox.Items.Clear();
            DestFieldComboBox.Items.Clear();
            foreach (string field in sourceList)
            {
                SourceFieldComboBox.Items.Add(field);
            }
            foreach (string field in targetList)
            {
                DestFieldComboBox.Items.Add(field);
            }

            List<object> tempList = (List<object>)finalFieldMap[(string)FieldTypesComboBox.SelectedValue];
            MappedListGrid.ItemsSource = tempList;
            MappedListGrid.Items.Refresh();

            foreach (object[] mappedField in tempList)
            {
                SourceFieldComboBox.Items.Remove((string)mappedField[0]);
                DestFieldComboBox.Items.Remove((string)mappedField[1]);
            }
        }

        private void SetFieldTypeList(ICollection keys)
        {
            if (finalFieldMap.Count == 0)
            {
                FieldTypesComboBox.Items.Clear();
                foreach (string key in keys)
                {
                    List<object> list = new List<object>();
                    finalFieldMap.Add(key, list);
                    FieldTypesComboBox.Items.Add(key);
                }
                FieldTypesComboBox.Items.Refresh();
            }
        }

        private void SetListsCopyFieldsTab(ICollection keys)
        {
            if (copyingFieldSet.Count == 0)
            {
                FieldTypes2ComboBox.Items.Clear();
                foreach (string key in keys)
                {
                    List<List<string>> fieldList = (List<List<string>>)fieldMap[key];
                    List<string> sourceList = (List<string>)fieldList.ElementAt(0);
                    List<object> list = new List<object>();
                    foreach (string value in sourceList)
                    {
                        object[] row = new object[2];
                        row[0] = value;
                        row[1] = false;
                        list.Add(row);
                    }
                    copyingFieldSet.Add(key, list);
                    FieldTypes2ComboBox.Items.Add(key);

                    object[] typeRow = new object[2];
                    typeRow[0] = key;
                    typeRow[1] = false;
                    migrateTypeSet.Add(typeRow);
                }
                FieldTypes2ComboBox.Items.Refresh();
            }

        }


        private void MapButton_Click(object sender, RoutedEventArgs e)
        {
            List<object> tempList = (List<object>)finalFieldMap[(string)FieldTypesComboBox.SelectedValue];
            object[] row = new object[3];
            row[0] = SourceFieldComboBox.SelectedValue;
            row[1] = DestFieldComboBox.SelectedValue;
            row[2] = false;
            tempList.Add(row);

            MappedListGrid.ItemsSource = tempList;
            MappedListGrid.Items.Refresh();

            SourceFieldComboBox.Items.Remove(SourceFieldComboBox.SelectedValue);
            DestFieldComboBox.Items.Remove(DestFieldComboBox.SelectedValue);
        }

        private void RemoveMapButton_Click(object sender, RoutedEventArgs e)
        {
            List<object> tempList = (List<object>)finalFieldMap[(string)FieldTypesComboBox.SelectedValue];
            foreach (object[] row in tempList.ToArray())
            {
                if ((bool)row[2])
                {
                    SourceFieldComboBox.Items.Add((string)row[0]);
                    DestFieldComboBox.Items.Add((string)row[1]);
                    tempList.Remove(row);
                }
            }
            MappedListGrid.ItemsSource = tempList;
            MappedListGrid.Items.Refresh();
        }

        private void NextButtonMapping_Click(object sender, RoutedEventArgs e)
        {
            StartTab.IsEnabled = true;
            StartTab.IsSelected = true;
        }

        private void FieldTypes2ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FieldTypes2ComboBox.SelectedValue != null)
            {
                List<object> tempList = (List<object>)copyingFieldSet[(string)FieldTypes2ComboBox.SelectedValue];
                FieldToCopyGrid.ItemsSource = tempList;
            }
        }

        private void NextButtonCopying_Click(object sender, RoutedEventArgs e)
        {
            FieldMappingTab.IsEnabled = true;
            FieldMappingTab.IsSelected = true;
            fieldMap = writeTarget.MapFields(readSource.workItemTypes);
        }

        private void CopyFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            writeTarget.SetFieldDefinitions(readSource.workItemTypes, copyingFieldSet);
            foreach (string key in copyingFieldSet.Keys)
            {
                List<object> list = (List<object>)copyingFieldSet[key];
                for (int i = 0; i < list.Count; i++)
                {
                    object[] field = (object[])list[i];
                    if ((bool)field[1])
                    {
                        list.Remove(field);
                        i--;
                    }
                } 
            }
            FieldToCopyGrid.Items.Refresh();
        }

        private void CopyWorkFlowsButton_Click(object sender, RoutedEventArgs e)
        {
            string error = writeTarget.ReplaceWorkFlow(readSource.workItemTypes, migrateTypeSet); 
            if (error.Length > 0)
            {
                System.Windows.MessageBox.Show(error, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            WorkFlowListGrid.Items.Refresh();
            //for (int i = 0; i < migrateTypeSet.Count; i++)
            //{
            //    object[] field = (object[])migrateTypeSet[i];
            //    if ((bool)field[1])
            //    {
            //        ItemCollection items = (ItemCollection)WorkFlowListGrid.Items.GetItemAt(i);
            //        foreach (ItemsControl item in items)
            //        {
            //            item.IsEnabled = false;
            //        }
            //    }
            //} 
        }

        private void CheckTestPlanHyperLink_Click(object sender, RoutedEventArgs e)
        {
            TestPlanViewUI ts = new TestPlanViewUI();
            ts.tfs = this.destinationTFS;
            ts.targetProjectName = this.destinationProject.Name;
            ts.printProjectName();
            ts.Show();
        }

        private void CheckLog_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"Log\Log-File");
        }

    }
}
