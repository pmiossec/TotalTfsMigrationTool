
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Xml;
using log4net;
using log4net.Config;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using MessageBox = System.Windows.MessageBox;
using TabControl = System.Windows.Controls.TabControl;

namespace TFSProjectMigration
{
    /// <summary>
    /// Interaction logic for TFSProjectMigrationUI.xaml
    /// </summary>
    public partial class TfsWorkItemMigrationUi
    {
        private TfsTeamProjectCollection _sourceTfs;
        private TfsTeamProjectCollection _destinationTfs;
        private WorkItemStore _sourceStore;
        private WorkItemStore _destinationStore;
        private Project _sourceProject;
        private Project _destinationProject;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(TfsWorkItemMigrationUi));
        public int MigrationState = 0;
        private bool _isNotIncludeClosed;
        private bool _isNotIncludeRemoved;
        private bool _areVersionHistoryCommentsIncluded;
        private bool _shouldWorkItemsBeLinkedToGitCommits;
        private WorkItemRead _readSource;
        private WorkItemWrite _writeTarget;
        private Hashtable _fieldMap;
        private Hashtable _finalFieldMap;
        private Hashtable _copyingFieldSet;
        private List<object> _migrateTypeSet;

        public TfsWorkItemMigrationUi()
        {
            InitializeComponent();
            XmlConfigurator.Configure();
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

                _finalFieldMap = new Hashtable();
                _copyingFieldSet = new Hashtable();
                _migrateTypeSet = new List<object>();
                _sourceTfs = tpp.SelectedTeamProjectCollection;
                _sourceStore = (WorkItemStore)_sourceTfs.GetService(typeof(WorkItemStore));

                _sourceProject = _sourceStore.Projects[tpp.SelectedProjects[0].Name];
                SourceProjectText.Text = string.Format("{0}/{1}", _sourceTfs.Uri, _sourceProject.Name);
                _readSource = new WorkItemRead(_sourceTfs, _sourceProject);

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
                _finalFieldMap = new Hashtable();
                _copyingFieldSet = new Hashtable();
                _migrateTypeSet = new List<object>();

                _destinationTfs = tpp.SelectedTeamProjectCollection;
                _destinationStore = (WorkItemStore)_destinationTfs.GetService(typeof(WorkItemStore));

                _destinationProject = _destinationStore.Projects[tpp.SelectedProjects[0].Name];
                DestinationProjectText.Text = string.Format("{0}/{1}", _destinationTfs.Uri, _destinationProject.Name);
                _writeTarget = new WorkItemWrite(_destinationTfs, _destinationProject);

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
                _writeTarget.CreateDefaultItemMapping(_readSource.WorkItemTypes);
                _isNotIncludeClosed = ClosedTextBox.IsChecked.GetValueOrDefault();
                _isNotIncludeRemoved = RemovedTextBox.IsChecked.GetValueOrDefault();
                _areVersionHistoryCommentsIncluded = VersionHistoryCheckBox.IsChecked.GetValueOrDefault();
                _shouldWorkItemsBeLinkedToGitCommits = LinkToCommitsCheckBox.IsChecked.GetValueOrDefault();
                ItemMappingTab.IsEnabled = true;
                ItemMappingTab.IsSelected = true;
            }
        }


        public void projectMigration()
        {
            Logger.InfoFormat("--------------------------------Migration from '{0}' to '{1}' Start----------------------------------------------", _sourceProject.Name, _destinationProject.Name);
            CheckTestPlanTextBlock.Dispatcher.BeginInvoke(new Action(delegate
            {
                CheckTestPlanTextBlock.Visibility = Visibility.Hidden;
            }));
            CheckLogTextBlock.Dispatcher.BeginInvoke(new Action(delegate
            {
                CheckLogTextBlock.Visibility = Visibility.Hidden;
            }));
            MigratingLabel.Dispatcher.BeginInvoke(new Action(delegate
            {
                MigratingLabel.Content = "Migrating...";
            }));

            StatusBar.Dispatcher.BeginInvoke(new Action(delegate
            {
                StatusBar.Visibility = Visibility.Visible;
            }));

            WorkItemCollection source = _readSource.GetWorkItems(_sourceProject.Name, _isNotIncludeClosed, _isNotIncludeRemoved, StatusBar, _writeTarget.GetMappedWorkItems()); //Get Workitems from source tfs 
            XmlNode[] iterations = _readSource.PopulateIterations(); //Get Iterations and Areas from source tfs 

            StatusViwer.Dispatcher.BeginInvoke(new Action(delegate
            {
                StatusViwer.Content = "Generating Areas...";
            }));
            _writeTarget.GenerateAreas(iterations[0], _sourceProject.Name); //Copy Areas

            StatusViwer.Dispatcher.BeginInvoke(new Action(delegate
            {
                StatusViwer.Content = StatusViwer.Content + "\nGenerating Iterations...";
            }));
            _writeTarget.GenerateIterations(iterations[1], _sourceProject.Name); //Copy Iterations

            StatusViwer.Dispatcher.BeginInvoke(new Action(delegate
            {
                StatusViwer.Content = StatusViwer.Content + "\nCopying Team Queries...";
            }));
            _writeTarget.SetTeamQueries(_readSource.QueryCol, _sourceProject.Name); //Copy Queries

            StatusViwer.Dispatcher.BeginInvoke(new Action(delegate
            {
                StatusViwer.Content = StatusViwer.Content + "\nCopying Work Items...";
            }));
            _writeTarget.writeWorkItems(_sourceStore, source, _sourceProject.Name, StatusBar, _finalFieldMap, _areVersionHistoryCommentsIncluded, _shouldWorkItemsBeLinkedToGitCommits); //Copy Workitems

            StatusViwer.Dispatcher.BeginInvoke(new Action(delegate
            {
                StatusViwer.Content = StatusViwer.Content + "\nCopying Test Plans...";
            }));
            TestPlanMigration tcm = new TestPlanMigration(_sourceTfs, _destinationTfs, _sourceProject.Name, _destinationProject.Name, _writeTarget.ItemMap, StatusBar);
            tcm.CopyTestPlans(); //Copy Test Plans

            MigratingLabel.Dispatcher.BeginInvoke(new Action(delegate
            {
                MigratingLabel.Content = "Project Migrated";
            }));

            StatusBar.Dispatcher.BeginInvoke(new Action(delegate
            {
                StatusBar.Visibility = Visibility.Hidden;
            }));
            CheckTestPlanTextBlock.Dispatcher.BeginInvoke(new Action(delegate
            {
                CheckTestPlanTextBlock.Visibility = Visibility.Visible;
            }));
            CheckLogTextBlock.Dispatcher.BeginInvoke(new Action(delegate
            {
                CheckLogTextBlock.Visibility = Visibility.Visible;
            }));
            Logger.Info("--------------------------------Migration END----------------------------------------------");
        }

        private void MigrationButton_Click(object sender, RoutedEventArgs e)
        {
            StartTab.IsEnabled = true;
            StartTab.IsSelected = true;
            StatusViwer.Content = "";
            MigratingLabel.Content = "";
            StatusBar.Value = 0;
            Thread migrationThread = new Thread(projectMigration);
            migrationThread.Start();
        }

        private void MigrationTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FieldMappingTab.IsSelected && e.OriginalSource is TabControl)
            {
                SetFieldTypeList(_fieldMap.Keys);
            }
            else if (FieldCopyTab.IsSelected && e.OriginalSource is TabControl)
            {
                SetListsCopyFieldsTab(_fieldMap.Keys);
                WorkFlowListGrid.ItemsSource = _migrateTypeSet;
                WorkFlowListGrid.Items.Refresh();
            }
            else if (ItemMappingTab.IsSelected && e.OriginalSource is TabControl)
            {
                SourceItemComboBox.Items.Clear();
                DestItemComboBox.Items.Clear();
                foreach (var sourceItem in _writeTarget.WorkItemTypeMap.Keys.OrderBy(k => k))
                {
                    SourceItemComboBox.Items.Add(sourceItem);
                }

                foreach (WorkItemType targetItem in _writeTarget.WorkItemTypes)
                {
                    DestItemComboBox.Items.Add(targetItem.Name);
                }

                DestItemComboBox.Items.Add("Do not Migrate");

                MappedItemListGrid.ItemsSource = _writeTarget.WorkItemTypeMap;
                MappedItemListGrid.Items.Refresh();
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
            List<List<string>> fieldList = (List<List<string>>)_fieldMap[fieldType];
            List<string> sourceList = fieldList.ElementAt(0);
            List<string> targetList = fieldList.ElementAt(1);

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

            List<object> tempList = (List<object>)_finalFieldMap[(string)FieldTypesComboBox.SelectedValue];
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
            if (_finalFieldMap.Count == 0)
            {
                FieldTypesComboBox.Items.Clear();
                foreach (string key in keys)
                {
                    List<object> list = new List<object>();
                    _finalFieldMap.Add(key, list);
                    FieldTypesComboBox.Items.Add(key);
                }
                FieldTypesComboBox.Items.Refresh();
            }
        }

        private void SetListsCopyFieldsTab(ICollection keys)
        {
            if (_copyingFieldSet.Count == 0)
            {
                FieldTypes2ComboBox.Items.Clear();
                foreach (string key in keys)
                {
                    List<List<string>> fieldList = (List<List<string>>)_fieldMap[key];
                    List<string> sourceList = fieldList.ElementAt(0);
                    List<object> list = new List<object>();
                    foreach (string value in sourceList)
                    {
                        object[] row = new object[2];
                        row[0] = value;
                        row[1] = false;
                        list.Add(row);
                    }
                    _copyingFieldSet.Add(key, list);
                    FieldTypes2ComboBox.Items.Add(key);

                    object[] typeRow = new object[2];
                    typeRow[0] = key;
                    typeRow[1] = false;
                    _migrateTypeSet.Add(typeRow);
                }
                FieldTypes2ComboBox.Items.Refresh();
            }
        }


        private void MapButton_Click(object sender, RoutedEventArgs e)
        {
            List<object> tempList = (List<object>)_finalFieldMap[(string)FieldTypesComboBox.SelectedValue];
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
            List<object> tempList = (List<object>)_finalFieldMap[(string)FieldTypesComboBox.SelectedValue];
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
                List<object> tempList = (List<object>)_copyingFieldSet[(string)FieldTypes2ComboBox.SelectedValue];
                FieldToCopyGrid.ItemsSource = tempList;
            }
        }

        private void NextButtonCopying_Click(object sender, RoutedEventArgs e)
        {
            FieldMappingTab.IsEnabled = true;
            FieldMappingTab.IsSelected = true;
            _fieldMap = _writeTarget.MapFields(_readSource.WorkItemTypes);
        }

        private void CopyFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            _writeTarget.SetFieldDefinitions(_readSource.WorkItemTypes, _copyingFieldSet);
            foreach (string key in _copyingFieldSet.Keys)
            {
                List<object> list = (List<object>)_copyingFieldSet[key];
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
            string error = _writeTarget.ReplaceWorkFlow(_readSource.WorkItemTypes, _migrateTypeSet);
            if (error.Length > 0)
            {
                MessageBox.Show(error, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            WorkFlowListGrid.Items.Refresh();
        }

        private void CheckTestPlanHyperLink_Click(object sender, RoutedEventArgs e)
        {
            TestPlanViewUI ts = new TestPlanViewUI();
            ts.tfs = _destinationTfs;
            ts.targetProjectName = _destinationProject.Name;
            ts.printProjectName();
            ts.Show();
        }

        private void CheckLog_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(@"Log\Log-File");
        }

        private void MapItemsButton_Click(object sender, RoutedEventArgs e)
        {
            if (SourceItemComboBox.SelectedItem == null || DestItemComboBox.SelectedItem == null)
                return;

            _writeTarget.WorkItemTypeMap[SourceItemComboBox.SelectedItem.ToString()] = DestItemComboBox.SelectedItem.ToString();

            MappedItemListGrid.Items.Refresh();

        }

        private void NextButtonItemMapping_Click(object sender, RoutedEventArgs e)
        {
            _fieldMap = _writeTarget.MapFields(_readSource.WorkItemTypes);
            FieldCopyTab.IsEnabled = true;
            FieldCopyTab.IsSelected = true;
        }

        private void SourceItemComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DestItemComboBox.SelectedItem = null;
            if (SourceItemComboBox.SelectedItem == null)
                return;

            var destItem = _writeTarget.WorkItemTypeMap[SourceItemComboBox.SelectedItem.ToString()];
            if (DestItemComboBox.Items.Contains(destItem))
                DestItemComboBox.SelectedItem = destItem;

        }
    }
}
