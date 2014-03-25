using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TDAPIOLELib;


using Ionic.Zip;
using System.IO;
using System.IO.Compression;

namespace ASAP
{
    public class QualityCenter
    {

        //***********************************List of Functions in Lib ******************************
        // 1.fAddTest
        // 2.fQCStepUpdate
        // 3.fUpdateTestStatusInQC
        // 4.fConnectToQC
        //***********************************List of Functions in Lib ******************************


        //*****************************************************************************************
        //*	Name		    : fAddTest
        //*	Description	    : Function to Add a QTP Test to Test Set
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string strTestName - Name of the test to be Added
        //*	Return Values	: Bool True on Success / False on failure
        //*****************************************************************************************
        public bool fAddTest(string strTestSetName, string strTestName)
        {

            //Check whether Node exist
            try
            {
                //Declare variables
                TreeManager objTM = (TreeManager)Global.objTD.TreeManager;
                TestFactory objTF = (TestFactory)Global.objTD.TestFactory;
                ITest test;

                //Get the test plan node
                SysTreeNode objAutoCoverageNode = (SysTreeNode)objTM.get_NodeByPath(Environment.GetEnvironmentVariable("TEST_PLAN_PATH"));
                string strTestFactoryFilter = "select TS_TEST_ID from TEST where TS_NAME = '" + strTestName + "' and TS_SUBJECT = " + objAutoCoverageNode.NodeID;

                //Fetch tests from Test factory corresponding to above defined filter
                List objTests = objTF.NewList(strTestFactoryFilter);
                int TestID;

                //Check Count
                if (objTests.Count == 0)
                {
                    //Set all Test Attributes and Post the test to DB
                    //object[] arrTestDetails = { strTestName, "QUICKTEST_TEST", Environment.UserName, (long)objAutoCoverageNode.NodeID };
                    //object[] arrTestDetails = { strTestName, "QUICKTEST_TEST", "ag3750", (long)objAutoCoverageNode.NodeID };
                    test = (ITest)objTF.AddItem(DBNull.Value);
                    test.Name = strTestName;
                    test["TS_SUBJECT"] = (long)objAutoCoverageNode.NodeID;
                    test["TS_TYPE"] = "QUICKTEST_TEST";
                    test.Post();
                    TestID = (int)test.ID;
                }
                else
                {
                    test = (ITest)objTests[1];
                    TestID = (int)test.ID;
                }

                Global.Dictionary.Add("TESTID", test.ID.ToString());

                //Declare Test Set Factory Object
                TestSetTreeManager objTSTM = (TestSetTreeManager)Global.objTD.TestSetTreeManager;
                TestSetFactory objTSF = (TestSetFactory)Global.objTD.TestSetFactory;

                //Get Test Set Tree Node
                SysTreeNode objTestSetFolder = (SysTreeNode)objTSTM.get_NodeByPath(Environment.GetEnvironmentVariable("TEST_SET_PATH"));
                string strTestSetFactoryFilter = "Select CY_CYCLE_ID from CYCLE where CY_CYCLE = '" + strTestSetName + "' And CY_FOLDER_ID = " + objTestSetFolder.NodeID;

                //Fetch Test Set from Test Set Factory
                List objTestSets = objTSF.NewList(strTestSetFactoryFilter);
                if (objTestSets.Count == 0)
                {
                    Global.fUpdateExecutionLog(LogType.error, "No Test Set by Name " + strTestSetName + " found in folder " + Environment.GetEnvironmentVariable("TEST_SET_PATH"));
                    return false;
                }

                //Declare Test Set Object
                TestSet objTS = (TestSet)objTestSets[1];
                TSTestFactory objTSTF = (TSTestFactory)objTS.TSTestFactory;

                //gET TEST sET id
                int TestSetID = (int)objTS.ID;
                Global.Dictionary.Add("TESTSETID", objTS.ID.ToString());
                Global.fUpdateExecutionLog(LogType.info, "Value of Test ID: " + TestID);

                TSTest objTST;

                //Check Test in TS Test Factory
                string strTSTestFactoryFilter = "Select * from TESTCYCL where TC_CYCLE_ID = " + TestSetID + " And TC_TEST_ID = " + TestID;
                List objTestSetTests = objTSTF.NewList(strTSTestFactoryFilter);
                if (objTestSetTests.Count == 0)
                {


                    //objTST = (TSTest)objTSTF.AddItem(TestID);
                    objTST = (TSTest)objTSTF.AddItem(DBNull.Value);
                    objTST["TC_TEST_ID"] = TestID;
                    objTST.Status = "Not Completed";
                    objTST.Post();
                }
                else
                {
                    objTST = (TSTest)objTestSetTests[1];
                }

                Global.Dictionary.Add("TSTESTID", objTST.ID.ToString());

                //Declare Run Factory variables
                RunFactory objRF = (RunFactory)objTST.RunFactory;
                Run objRun = (Run)objRF.AddItem(DBNull.Value);
                objRun["RN_TEST_ID"] = TestID;
                objRun.Name = strTestName;
                objRun.Status = "Not Completed";
                objRun.Post();

                Global.Dictionary.Add("RUNID", objRun.ID.ToString());
                Global.fUpdateExecutionLog(LogType.info, "RUN ID :   " + objRun.ID.ToString());


                return true;
            }

            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Got exception while executing the fAddTest function. Exception + " + e + " occured");
                return false;
            }

        }


        //*****************************************************************************************
        //*	Name		    : fQCStepUpdate
        //*	Description	    : Function to Add a Step to QC Test in Test Set
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string strStepName - Name of the Step
        //*                 : string strStepDesc - Description of the Step
        //*                 : string strExpValue - Expected Value
        //*                 : string strActValue - Actual value
        //*                 : string strResult   - Result
        //*	Return Values	: Bool True on Success / False on failure
        //*****************************************************************************************
        public bool fQCStepUpdate(string strStepName, string strStepDesc, string strActValue, string strResult)
        {

            try
            {
                //Declare RUn factory variables
                RunFactory objRF = (RunFactory)Global.objTD.RunFactory;

                //CreateFilter
                string strRunFactoryFilter = "Select * from RUN where RN_RUN_ID = " + Global.Dictionary["RUNID"];
                List objRuns;

                try
                {
                    //Create New List
                    Global.fUpdateExecutionLog(LogType.info, "RUN QUERY  :   " + strRunFactoryFilter);
                    objRuns = objRF.NewList(strRunFactoryFilter);
                }
                catch (Exception e)
                {

                    TDFilter TDRunFactoryFilter = (TDFilter)objRF.Filter;
                    TDRunFactoryFilter["RN_RUN_ID"] = Global.Dictionary["RUNID"];
                    Global.fUpdateExecutionLog(LogType.info, "RUN QUERY  :   " + TDRunFactoryFilter.Text);
                    objRuns = objRF.NewList(TDRunFactoryFilter.Text);
                }

                //Global.fUpdateExecutionLog(LogType.info, "RUN QUERY  :   " + strRunFactoryFilter);


                //Check Count
                if (objRuns.Count == 0)
                {
                    return false;
                }

                Run objRun = (Run)objRuns[1];

                //Set Step factory variables
                StepFactory objSF = (StepFactory)objRun.StepFactory;
                Step objStep = (Step)objSF.AddItem(DBNull.Value);
                objStep["ST_STEP_NAME"] = strStepName;
                objStep["ST_DESCRIPTION"] = strStepDesc;
                objStep["ST_STATUS"] = strResult;
                objStep["ST_ACTUAL"] = strActValue;
                objStep.Post();
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured while updating step in QC");
                return false;
            }

            return true;
        }

        //*****************************************************************************************
        //*	Name		    : fUpdateTestStatusInQC
        //*	Description	    : Updates the status of Current RUn and Overall execution for a particular test
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string strStatus - Test Status
        //*	Return Values	: Bool True on Success / False on failure
        //*****************************************************************************************
        public bool fUpdateTestStatusInQC(string strStatus)
        {
            try
            {

                //Declare RUn factory variables
                RunFactory objRF = (RunFactory)Global.objTD.RunFactory;

                //CreateFilter
                string strRunFactoryFilter = "Select * from RUN where RN_RUN_ID = " + Global.Dictionary["RUNID"];
                List objRuns;

                try
                {
                    //Create New List
                    Global.fUpdateExecutionLog(LogType.info, "RUN QUERY  :   " + strRunFactoryFilter);
                    objRuns = objRF.NewList(strRunFactoryFilter);
                }
                catch (Exception e)
                {

                    TDFilter TDRunFactoryFilter = (TDFilter)objRF.Filter;
                    TDRunFactoryFilter["RN_RUN_ID"] = Global.Dictionary["RUNID"];
                    Global.fUpdateExecutionLog(LogType.info, "RUN QUERY  :   " + TDRunFactoryFilter.Text);
                    objRuns = objRF.NewList(TDRunFactoryFilter.Text);
                }

                //Check Count
                if (objRuns.Count == 0) return false;

                Run objRun = (Run)objRuns[1];
                objRun.Status = strStatus;
                objRun.Post();
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured while updating test status in QC");
                return false;
            }

            return true;
        }

        //*****************************************************************************************
        //*	Name		    : fConnectToQC
        //*	Description	    : Makes a connection with QC using Specified credentials
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string strStatus - Test Status
        //*	Return Values	: Bool True on Success / False on failure
        //*****************************************************************************************
        public TDAPIOLELib.TDConnection fConnectToQC(string strQCServer, string strQCUser, string strQCPassword, string strQCDomain, string strQCProject)
        {

            TDAPIOLELib.TDConnection objTD = new TDAPIOLELib.TDConnection();

            try
            {
                objTD.InitConnectionEx(strQCServer);
                objTD.Login(strQCUser, strQCPassword);
                objTD.Connect(strQCDomain, strQCProject);

                //Check if Connection is successfull
                if (objTD.Connected != true)
                {
                    Global.fUpdateExecutionLog(LogType.error, "Unable to connect to QC");
                    return objTD;

                }
                else
                {
                    return objTD;

                }
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Unable to Connnect to QC. Error Message : " + e.Message);
                return null;

            }

        }

        //*****************************************************************************************
        //*	Name		    : fAttachResultsToRun
        //*	Description	    : Attach the HTML Reports and the Screenshots to Current RUN
        //*	Author		    : Aniket Gadre
        //*	Input Params	: strTestDetails
        //*	Return Values	: Bool True on Success / False on failure
        //*****************************************************************************************
        public bool fAttachResultsToRun(string strTestDetails)
        {
            //Set the name for the Test Case Report File
            string strHTMLReports = Environment.GetEnvironmentVariable("HTML_REPORTS_PATH");
            string strHTMLReport = strHTMLReports + "Report_" + strTestDetails + ".html";
            string strSnapShotsFolder = Environment.GetEnvironmentVariable("SCREEN_SHOT_PATH") + strTestDetails;
            string strZipFilePath = strHTMLReports + "Report_" + strTestDetails + ".zip";

            try
            {

                //Create Zip Directory
                string zipDir = strHTMLReports + "Reports";
                if (Directory.Exists(zipDir)) Directory.Delete(zipDir, true);
                Directory.CreateDirectory(zipDir + "\\Screen_Prints\\" + strTestDetails);

                //COpy required dir into zip dir
                File.Copy(strHTMLReport, zipDir + "\\Report_" + strTestDetails + ".html");
                fCopyDirectory(new DirectoryInfo(strSnapShotsFolder), new DirectoryInfo(zipDir + "\\Screen_Prints\\" + strTestDetails));

                //Zip DIR
                //Create a zip FIle
                ZipFile objZip = new ZipFile();

                //Add required File and Folder
                objZip.AddDirectory(zipDir);
                objZip.Save(strZipFilePath);


                //Declare RUn factory variables
                RunFactory objRF = (RunFactory)Global.objTD.RunFactory;

                //CreateFilter
                string strRunFactoryFilter = "Select * from RUN where RN_RUN_ID = " + Global.Dictionary["RUNID"];
                List objRuns;

                try
                {
                    //Create New List
                    Global.fUpdateExecutionLog(LogType.info, "RUN QUERY  :   " + strRunFactoryFilter);
                    objRuns = objRF.NewList(strRunFactoryFilter);
                }
                catch (Exception e)
                {

                    TDFilter TDRunFactoryFilter = (TDFilter)objRF.Filter;
                    TDRunFactoryFilter["RN_RUN_ID"] = Global.Dictionary["RUNID"];
                    Global.fUpdateExecutionLog(LogType.info, "RUN QUERY  :   " + TDRunFactoryFilter.Text);
                    objRuns = objRF.NewList(TDRunFactoryFilter.Text);
                }

                //Check Count
                if (objRuns.Count == 0) return false;

                //Get Current Run
                Run objRun = (Run)objRuns[1];

                //Get Attachement Factory Object object
                AttachmentFactory objAF = (AttachmentFactory)objRun.Attachments;
                Attachment objA = (Attachment)objAF.AddItem(DBNull.Value);
                objA.FileName = strZipFilePath;
                objA.Type = 1;
                objA.Post();

                //Declare the variable for Test Set Factory
                TestSetFactory objTSF = (TestSetFactory)Global.objTD.TestSetFactory;
                string strTestSetFactoryFilter = "Select * from CYCLE where CY_CYCLE_ID = " + Global.Dictionary["TESTSETID"];

                //Fetch Test Set from Test Set Factory
                List objTestSets = objTSF.NewList(strTestSetFactoryFilter);

                //Declare Test Set Object
                TestSet objTS = (TestSet)objTestSets[1];
                TSTestFactory objTSTF = (TSTestFactory)objTS.TSTestFactory;

                //Check Test in TS Test Factory
                string strTSTestFactoryFilter = "Select * from TESTCYCL where TC_CYCLE_ID = " + Global.Dictionary["TESTSETID"] + " And TC_TEST_ID = " + Global.Dictionary["TESTID"];
                List objTestSetTests = objTSTF.NewList(strTSTestFactoryFilter);
                TSTest objTST = (TSTest)objTestSetTests[1];

                ////Get Attachment Factory Object
                objAF = (AttachmentFactory)objTST.Attachments;
                objA = (Attachment)objAF.AddItem(DBNull.Value);
                objA.FileName = strZipFilePath;
                objA.Type = 1;
                objA.Post();

                //Delete the zip file
                File.Delete(strZipFilePath);
                Directory.Delete(zipDir, true);
                Global.fUpdateExecutionLog(LogType.info, "Attaching results successfull");
                return true;

            }

            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Attaching results in QC Failed. Error Message: " + e);
                return false;
            }


        }


        //*****************************************************************************************
        //*	Name		    : fCopyDirectory
        //*	Description	    : Copy from one dir to another
        //*	Author		    : Aniket Gadre
        //*	Input Params	: Source, target
        //*	Return Values	: Bool True on Success / False on failure
        //*****************************************************************************************
        private bool fCopyDirectory(DirectoryInfo source, DirectoryInfo destination)
        {
            try
            {

                // Copy all files.
                FileInfo[] files = source.GetFiles();
                foreach (FileInfo file in files)
                {
                    file.CopyTo(Path.Combine(destination.FullName,
                        file.Name));
                }

                // Process subdirectories.
                DirectoryInfo[] dirs = source.GetDirectories();
                foreach (DirectoryInfo dir in dirs)
                {
                    // Get destination directory.
                    string destinationDir = Path.Combine(destination.FullName, dir.Name);

                    // Call CopyDirectory() recursively.
                    fCopyDirectory(dir, new DirectoryInfo(destinationDir));
                }
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Copying SS from source to destination failed. Error Message: " + e);
                return false;
            }

            return true;
        }
    }
}
