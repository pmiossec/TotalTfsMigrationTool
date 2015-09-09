using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.Server;
using log4net.Config;
using log4net;
using System.Windows.Controls;

namespace TFSProjectMigration
{
    public class TestPlanMigration
    {
        ITestManagementTeamProject sourceproj;
        ITestManagementTeamProject destinationproj;
        public Hashtable workItemMap;
        ProgressBar progressBar;
        String projectName;
        private static readonly ILog logger = LogManager.GetLogger(typeof(TFSWorkItemMigrationUI));

        public TestPlanMigration(TfsTeamProjectCollection sourceTfs, TfsTeamProjectCollection destinationTfs, string sourceProject, string destinationProject, Hashtable workItemMap, ProgressBar progressBar)
        {
            this.sourceproj = GetProject(sourceTfs, sourceProject);
            this.destinationproj = GetProject(destinationTfs, destinationProject);
            this.workItemMap = workItemMap;
            this.progressBar = progressBar;
            projectName = sourceProject;
        }

        private ITestManagementTeamProject GetProject(TfsTeamProjectCollection tfs, string project)
        {
            
            ITestManagementService tms = tfs.GetService<ITestManagementService>();

            return tms.GetTeamProject(project);
        }
        public void CopyTestPlans()
        {
            int i = 1;
            int planCount= sourceproj.TestPlans.Query("Select * From TestPlan").Count;
            //delete Test Plans if any existing test plans.
            //foreach (ITestPlan destinationplan in destinationproj.TestPlans.Query("Select * From TestPlan"))
            //{

            //    System.Diagnostics.Debug.WriteLine("Deleting Plan - {0} : {1}", destinationplan.Id, destinationplan.Name);

            //    destinationplan.Delete(DeleteAction.ForceDeletion); ;

            //}
           
            foreach (ITestPlan sourceplan in sourceproj.TestPlans.Query("Select * From TestPlan"))
            {
                System.Diagnostics.Debug.WriteLine("Plan - {0} : {1}", sourceplan.Id, sourceplan.Name);

                ITestPlan destinationplan = destinationproj.TestPlans.Create();

                destinationplan.Name = sourceplan.Name;
                destinationplan.Description = sourceplan.Description;
                destinationplan.StartDate = sourceplan.StartDate;
                destinationplan.EndDate = sourceplan.EndDate;
                destinationplan.State = sourceplan.State;            
                destinationplan.Save();

                //drill down to root test suites.
                if (sourceplan.RootSuite != null && sourceplan.RootSuite.Entries.Count > 0)
                {
                    CopyTestSuites(sourceplan, destinationplan);
                }

                destinationplan.Save();

                progressBar.Dispatcher.BeginInvoke(new Action(delegate()
                {
                    float progress = (float)i / (float) planCount;

                    progressBar.Value = ((float)i / (float) planCount) * 100;
                }));
                i++;
            }

        }

        //Copy all Test suites from source plan to destination plan.
        private void CopyTestSuites(ITestPlan sourceplan, ITestPlan destinationplan)
        {
            ITestSuiteEntryCollection suites = sourceplan.RootSuite.Entries;
            CopyTestCases(sourceplan.RootSuite, destinationplan.RootSuite);

            foreach (ITestSuiteEntry suite_entry in suites)
            {
                IStaticTestSuite suite = suite_entry.TestSuite as IStaticTestSuite;
                if (suite != null)
                {
                    IStaticTestSuite newSuite = destinationproj.TestSuites.CreateStatic();
                    newSuite.Title = suite.Title;
                    destinationplan.RootSuite.Entries.Add(newSuite);
                    destinationplan.Save();

                    CopyTestCases(suite, newSuite);
                    if (suite.Entries.Count > 0)
                        CopySubTestSuites(suite, newSuite);
                }
            }

        }
        //Drill down and Copy all subTest suites from source root test suite to destination plan's root test suites.
        private void CopySubTestSuites(IStaticTestSuite parentsourceSuite, IStaticTestSuite parentdestinationSuite)
        {
            ITestSuiteEntryCollection suitcollection = parentsourceSuite.Entries;
            foreach (ITestSuiteEntry suite_entry in suitcollection)
            {
                IStaticTestSuite suite = suite_entry.TestSuite as IStaticTestSuite;
                if (suite != null)
                {
                    IStaticTestSuite subSuite = destinationproj.TestSuites.CreateStatic();
                    subSuite.Title = suite.Title;
                    parentdestinationSuite.Entries.Add(subSuite);

                    CopyTestCases(suite, subSuite);

                    if (suite.Entries.Count > 0)
                        CopySubTestSuites(suite, subSuite);

                }
            }


        }

        //Copy all subTest suites from source root test suite to destination plan's root test suites.
        private void CopyTestCases(IStaticTestSuite sourcesuite, IStaticTestSuite destinationsuite)
        {

            ITestSuiteEntryCollection suiteentrys = sourcesuite.TestCases;

            foreach (ITestSuiteEntry testcase in suiteentrys)
            {
                try
                {   //check whether testcase exists in new work items(closed work items may not be created again).
                    if (!workItemMap.ContainsKey(testcase.TestCase.WorkItem.Id))
                    {
                        continue;
                    }

                    int newWorkItemID = (int)workItemMap[testcase.TestCase.WorkItem.Id];
                    ITestCase tc = destinationproj.TestCases.Find(newWorkItemID);
                    destinationsuite.Entries.Add(tc);

                    bool updateTestCase = false;
                    TestActionCollection testActionCollection = tc.Actions;
                    foreach (var item in testActionCollection)
                    {
                        var sharedStepRef = item as ISharedStepReference;
                        if (sharedStepRef != null)
                        {

                            int newSharedStepId = (int)workItemMap[sharedStepRef.SharedStepId];
                            //GetNewSharedStepId(testCase.Id, sharedStepRef.SharedStepId);
                            if (0 != newSharedStepId)
                            {
                                sharedStepRef.SharedStepId = newSharedStepId;
                                updateTestCase = true;
                            }

                        }
                    }
                    if (updateTestCase)
                    {
                        Console.WriteLine();
                        Console.WriteLine("Test case with Id: {0} updated", tc.Id);
                        tc.Save();
                    }
                }
                catch (Exception)
                {
                    logger.Info("Error retrieving Test case  " + testcase.TestCase.WorkItem.Id + ": " + testcase.Title);
                }
            }
        }

    }


    
}
