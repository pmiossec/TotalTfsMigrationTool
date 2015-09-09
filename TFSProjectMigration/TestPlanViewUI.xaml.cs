using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using System.ComponentModel;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TFSProjectMigration
{
    /// <summary>
    /// Interaction logic for TestPlanViewUI.xaml
    /// </summary>
    public partial class TestPlanViewUI : Window
    {
        TfsTeamProjectCollection _tfs;
        WorkItemStore _store;

        public TfsTeamProjectCollection tfs
        {
            get { return _tfs; }
            set { _tfs = value; }
        }
        string projectName = "aa";
        Window mainWindow;
        public string targetProjectName
        {
            get { return projectName; }
            set { projectName = value; }
        }

        public TestPlanViewUI()
        {
            InitializeComponent();
        }

        public void printProjectName()
        {
            Labal1.Content = projectName;
            ViewTestPlans();
        }




        private void ViewTestPlans()
        {
            if (_tfs != null)
            {
                ITestManagementService test_service = (ITestManagementService)_tfs.GetService(typeof(ITestManagementService));
                _store = (WorkItemStore)_tfs.GetService(typeof(WorkItemStore));
                ITestManagementTeamProject _testproject = test_service.GetTeamProject(projectName);
                GetTestPlans(_testproject);
            }
            else
            {
                Console.WriteLine("  aaa ");
            }
        }


        void GetTestPlans(ITestManagementTeamProject testproject)
        {
            ITestPlanCollection plans = testproject.TestPlans.Query("Select * From TestPlan");

            TreeViewItem root = null;
            root = new TreeViewItem();
            root.Header = ImageHelpers.CreateHeader(testproject.WitProject.Name, ItemTypes.TeamProject);
            TreeMain.Items.Add(root);

            foreach (ITestPlan plan in plans)
            {
                TreeViewItem plan_tree = new TreeViewItem();
                plan_tree.Header = ImageHelpers.CreateHeader(plan.Name, ItemTypes.TestPlan);

                if (plan.RootSuite != null && plan.RootSuite.Entries.Count > 0)
                    GetPlanSuites(plan.RootSuite.Entries, plan_tree);

                root.Items.Add(plan_tree);
            }
        }

        void GetPlanSuites(ITestSuiteEntryCollection suites, TreeViewItem tree_item)
        {
            foreach (ITestSuiteEntry suite_entry in suites)
            {
                IStaticTestSuite suite = suite_entry.TestSuite as IStaticTestSuite;
                if (suite != null)
                {
                    TreeViewItem suite_tree = new TreeViewItem();
                    suite_tree.Header = ImageHelpers.CreateHeader(suite.Title, ItemTypes.TestSuite);

                    GetTestCases(suite, suite_tree);

                    tree_item.Items.Add(suite_tree);

                    if (suite.Entries.Count > 0)
                        GetPlanSuites(suite.Entries, suite_tree);

                }
            }
        }

        void GetTestCases(IStaticTestSuite suite, TreeViewItem tree_item)
        {
            //AllTestCases - Will show all the Test Cases under that Suite even in sub suites.
            //ITestCaseCollection testcases = suite.AllTestCases;

            //Will bring only the Test Case under a specific Test Suite.
            ITestSuiteEntryCollection suiteentrys = suite.TestCases;

            foreach (ITestSuiteEntry testcase in suiteentrys)
            {
                TreeViewItem test = new TreeViewItem();
                test.Header = ImageHelpers.CreateHeader(testcase.Title, ItemTypes.TestCase);
                tree_item.Items.Add(test);
            }
        }
    }
}
