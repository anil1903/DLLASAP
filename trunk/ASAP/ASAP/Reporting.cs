using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;



namespace ASAP
{
    public class Reporting
    {
        // Filenames
        private string g_strTestCaseReport;
        private string g_strSnapshotFolderName;
        private string g_strScriptName;

        //Counters and Integers
        private int g_iSnapshotCount;
        private int g_OperationCount;
        private int g_iPassCount;
        private int g_iFailCount;
        private int g_iTCPassed;
        private int g_iTestCaseNo;

        //Timers
        private DateTime g_StartTime;
        private DateTime g_EndTime;
        private DateTime g_SummaryStartTime;
        private DateTime g_SummaryEndTime;


        //Global Dictionary
        // private Dictionary<string, string> Dictionary = new Dictionary<string, string>();



        ////***************************************************************************************** 
        ////*     Name            : fnSetReporterParams 
        ////*     Description     : 
        ////*     Author          :  Aniket Gadre 
        ////*     Input Params    :       None 
        ////*     Return Values   :       None 
        ////***************************************************************************************** 
        //public void fnSetReporterParams(Dictionary<string, string> Dict, String strReportsPath, TDAPIOLELib.TDConnection TD)
        //{
        //    Dictionary = Dict;
        //    Environment.GetEnvironmentVariable("HTML_REPORTS_PATH") = strReportsPath;
        //    objTD = TD;
        //}




        //*****************************************************************************************
        //*	Name		: fnCreateSummaryReport
        //*	Description	: The function creates the summary HTML file
        //*	Author		:  Aniket Gadre
        //*	Input Params	: 	None
        //*	Return Values	: 	None

        //*****************************************************************************************
        public void fnCreateSummaryReport()
        {
            //Setting counter value
            g_iTCPassed = 0;
            g_iTestCaseNo = 0;
            g_SummaryStartTime = DateTime.Now;

            //Open the test case report for writing	               
            StreamWriter fileStream = new StreamWriter(Environment.GetEnvironmentVariable("HTML_REPORTS_PATH") + "SummaryReport.html", true);

            //Write the initial comments into the file
            fileStream.WriteLine("<HTML><BODY><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=100% BGCOLOR=BLACK>");
            fileStream.WriteLine("<TR><TD WIDTH=90% ALIGN=CENTER BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=ORANGE SIZE=3><B>AMDOCS</B></FONT></TD></TR><TR><TD ALIGN=CENTER BGCOLOR=ORANGE><FONT FACE=VERDANA COLOR=WHITE SIZE=3><B>Selenium Framework Reporting</B></FONT></TD></TR></TABLE><TABLE CELLPADDING=3 WIDTH=100%><TR height=30><TD WIDTH=100% ALIGN=CENTER BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=//0073C5 SIZE=2><B>&nbsp; Automation Result : " + DateTime.Now + " on Machine " + Environment.MachineName + " by user " + Environment.UserName + " on Environment " + Environment.GetEnvironmentVariable("ENV_KEY") + " </B></FONT></TD></TR><TR HEIGHT=5></TR></TABLE>");
            fileStream.WriteLine("<TABLE  CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
            fileStream.WriteLine("<TR COLS=6 BGCOLOR=ORANGE><TD WIDTH=10%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>TC No.</B></FONT></TD><TD  WIDTH=70%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Test Name</B></FONT></TD><TD BGCOLOR=ORANGE WIDTH=30%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Status</B></FONT></TD></TR>");

            //Close the file
            fileStream.Flush();
            fileStream.Close();
            fileStream = null;

        }

        //*****************************************************************************************
        //*	Name		    : fnCreateHtmlReport
        //*	Description	    : The function creates the result HTML file
        //*	                  In Case the file already exists, it will overwrite it and also delete the existing folders.
        //*	Author		    : Aniket Gadre
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public void fnCreateHtmlReport(string strTestName)
        {

            //Set the default Operation count as 0
            g_OperationCount = 0;

            //Number of default Pass and Fail cases to 0
            g_iPassCount = 0;
            g_iFailCount = 0;

            //Snapshot count to start from 0
            g_iSnapshotCount = 0;

            //script name
            g_strScriptName = strTestName;

            //Set the name for the Test Case Report File
            g_strTestCaseReport = Environment.GetEnvironmentVariable("HTML_REPORTS_PATH") + "Report_" + g_strScriptName + ".html";

            //Snap Shot folder
            g_strSnapshotFolderName = Environment.GetEnvironmentVariable("SCREEN_SHOT_PATH") + g_strScriptName;

            //Create the folder for snapshots
            if (Directory.Exists(g_strSnapshotFolderName))
            {
                Directory.Delete(g_strSnapshotFolderName, true);
            }
            Directory.CreateDirectory(g_strSnapshotFolderName);

            //Open the HTML file
            StreamWriter fileStream = new StreamWriter(g_strTestCaseReport, true);

            //Write the Test Case name and allied headers into the file
            fileStream.WriteLine("<HTML><BODY><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=100% BGCOLOR=ORANGE>");
            fileStream.WriteLine("<TR><TD WIDTH=90% ALIGN=CENTER BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=ORANGE SIZE=3><B>AMDOCS</B></FONT></TD></TR><TR><TD ALIGN=CENTER BGCOLOR=ORANGE><FONT FACE=VERDANA COLOR=WHITE SIZE=3><B>ASAP Framework Reporting</B></FONT></TD></TR></TABLE><TABLE CELLPADDING=3 WIDTH=100%><TR height=30><TD WIDTH=100% ALIGN=CENTER BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=//0073C5 SIZE=2><B>&nbsp; Automation Result : " + DateTime.Now + " on Machine " + Environment.MachineName + " by user " + Environment.UserName + "</B></FONT></TD></TR><TR HEIGHT=5></TR></TABLE>");
            fileStream.WriteLine("<TABLE BORDER=0 BORDERCOLOR=WHITE CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
            fileStream.WriteLine("<TR><TD BGCOLOR=BLACK WIDTH=20%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Test	 Name:</B></FONT></TD><TD COLSPAN=6 BGCOLOR=BLACK><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>" + g_strScriptName + "</B></FONT></TD></TR>");
            //fileStream.WriteLine("<TR><TD BGCOLOR=BLACK WIDTH=20%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Test	Iteration:</B></FONT></TD><TD COLSPAN=6 BGCOLOR=BLACK><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B> </B></FONT></TD></TR>");
            fileStream.WriteLine("</TABLE><BR/><TABLE WIDTH=100% CELLPADDING=3>");
            fileStream.WriteLine("<TR WIDTH=100%><TH BGCOLOR=ORANGE WIDTH=10%><FONT FACE=VERDANA SIZE=2>Step No.</FONT></TH><TH BGCOLOR=ORANGE WIDTH=30%><FONT FACE=VERDANA SIZE=2>Step Description</FONT></TH><TH BGCOLOR=ORANGE WIDTH=30%><FONT FACE=VERDANA SIZE=2>Obtained Value</FONT></TH><TH BGCOLOR=ORANGE WIDTH=20%><FONT FACE=VERDANA SIZE=2>Result</FONT></TH></TR>");
            fileStream.Flush();
            fileStream.Close();

            //start time
            g_StartTime = DateTime.Now;
        }
        //*****************************************************************************************
        //*	End of function fnCreateHtmlReport
        //*****************************************************************************************

        //*****************************************************************************************
        //*	Name		: fnWriteTestSummary
        //*	Description	: The function Writes the final outcome of a test case to a summary file.
        //*	Author		:  Aniket Gadre
        //*	Input Params	: 	
        //*			strTestCaseName(String) - the name of the test case
        //*			strResult(String) - the result (Pass/Fail)
        //*	Return Values	: 	
        //*			(Boolean) TRUE - Succeessful write
        //*				 FALSE - Report file not created
        //*****************************************************************************************
        public void fnWriteTestSummary(string strTestCaseName, string strResult)
        {

            string sColor, sRowColor;

            //Open the Test Summary Report File
            StreamWriter fileStream = new StreamWriter(Environment.GetEnvironmentVariable("HTML_REPORTS_PATH") + "SummaryReport.html", true);

            //Check color result
            if (strResult.ToUpper() == "PASSED" || strResult.ToUpper() == "PASS")
            {
                sColor = "GREEN";
                g_iTCPassed++;
            }
            else if (strResult.ToUpper() == "FAILED" || strResult.ToUpper() == "FAIL")
            {
                sColor = "RED";
            }
            else
            {
                sColor = "ORANGE";
            }

            g_iTestCaseNo++;

            if (g_iTestCaseNo % 2 == 0)
            {
                //sRowColor = "//BEBEBE";
                sRowColor = "#EEEEEE";
            }
            else
            {
                sRowColor = "#D3D3D3";
            }

            //Write the result of Individual Test Case
            fileStream.WriteLine("<TR COLS=3 BGCOLOR=" + sRowColor + "><TD  WIDTH=10%><FONT FACE=VERDANA SIZE=2>" + g_iTestCaseNo + "</FONT></TD><TD  WIDTH=70%><FONT FACE=VERDANA SIZE=2>" + strTestCaseName + "</FONT></TD><TD  WIDTH=20%><A HREF='" + strTestCaseName + ".html'><FONT FACE=VERDANA SIZE=2 COLOR=" + sColor + "><B>" + strResult + "</B></FONT></A></TD></TR>");

            //Close the file
            fileStream.Flush();
            fileStream.Close();
            fileStream = null;

        }

        //*****************************************************************************************
        //*	Name		: fnCloseHtmlReport
        //*	Description	: The function Closes the HTML file
        //*	Author		: Aniket Gadre
        //*	Input Params	: 	None
        //*	Return Values	: 	None
        //*****************************************************************************************
        public void fnCloseHtmlReport()
        {

            //Declarations
            string strTestCaseResult;
            double dblTotalTimeTaken;
            string g_strTestCaseReport = Environment.GetEnvironmentVariable("HTML_REPORTS_PATH") + "Report_" + g_strScriptName + ".html";

            //Open the HTML file
            StreamWriter fileStream = new StreamWriter(g_strTestCaseReport, true);

            //Get the Current time
            g_EndTime = DateTime.Now;

            //Get Time Taken in Minutes & Seconds
            dblTotalTimeTaken = Math.Round((g_EndTime - g_StartTime).TotalMinutes, 2);
            int TotalMins = (int)dblTotalTimeTaken;
            int RemainingSeconds = (int)((dblTotalTimeTaken - TotalMins) * 60);

            //Write the number of test steps passed/failed and the time which the test case took to run
            fileStream.WriteLine("<TR></TR><TR><TD BGCOLOR=BLACK WIDTH=10%></TD><TD BGCOLOR=BLACK WIDTH=30%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Time Taken : " + TotalMins + ":" + RemainingSeconds + "</B></FONT></TD><TD BGCOLOR=BLACK WIDTH=30%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Pass Count : " + g_iPassCount + "</B></FONT></TD><TD BGCOLOR=BLACK WIDTH=20%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Fail Count : " + g_iFailCount + "</b></FONT></TD></TR>");
            //Close the HTML tags
            fileStream.WriteLine("</TABLE><TABLE WIDTH=100%><TR><TD ALIGN=RIGHT><FONT FACE=VERDANA COLOR=ORANGE SIZE=1>&copy; amdocs - Integrated Customer Management</FONT></TD></TR></TABLE></BODY></HTML>");
            //Close the HTML report
            fileStream.Flush();
            fileStream.Close();


            //Check if Test Case has passed or failed
            if (g_iFailCount != 0)
            {
                strTestCaseResult = "Failed";
            }
            else
            {
                strTestCaseResult = "Passed";
            }

            //Call fUpdateTestStatusInQC

            if (Global.objTD != null && Global.objTD.Connected == true)
            {
                QualityCenter objQC = new QualityCenter();
                objQC.fUpdateTestStatusInQC(strTestCaseResult);
            }

            //Write into the Summary Report
            fnWriteTestSummary("Report_" + g_strScriptName, strTestCaseResult);

        }
        //*****************************************************************************************
        //*	End of function fnCloseHtmlReport
        //*****************************************************************************************

        //*****************************************************************************************
        //*	Name		: fnCloseTestSummary
        //*	Description	: The function Closes the summary file
        //*	Author		:  Aniket Gadre
        //*	Input Params	: 	None
        //*	Return Values	: 	None
        //*****************************************************************************************
        public void fnCloseTestSummary()
        {
            g_SummaryEndTime = DateTime.Now;

            //Open the Test Summary Report File
            StreamWriter fileStream = new StreamWriter(Environment.GetEnvironmentVariable("HTML_REPORTS_PATH") + "SummaryReport.html", true);

            //Calculate total time taken
            double dblTotalTimeTaken = Math.Round((g_SummaryEndTime - g_SummaryStartTime).TotalMinutes, 2);
            int TotalMins = (int)dblTotalTimeTaken;
            int RemainingSeconds = (int)((dblTotalTimeTaken - TotalMins) * 60);

            //Close the HTML tags
            fileStream.WriteLine("</TABLE><TABLE WIDTH=100%><TR>");
            fileStream.WriteLine("<TD BGCOLOR=BLACK WIDTH=10%></TD><TD BGCOLOR=BLACK WIDTH=70%><FONT FACE=VERDANA SIZE=2 COLOR=WHITE><B>Total Execution Time : " + TotalMins + ":" + RemainingSeconds + "</B></FONT></TD><TD BGCOLOR=BLACK WIDTH=20%><FONT FACE=WINGDINGS SIZE=4>2</FONT><FONT FACE=VERDANA SIZE=2 COLOR=WHITE><B>Total Passed: " + g_iTCPassed + "</B></FONT></TD>");
            fileStream.WriteLine("</TR></TABLE>");
            fileStream.WriteLine("<TABLE WIDTH=100%><TR><TD ALIGN=RIGHT><FONT FACE=VERDANA COLOR=ORANGE SIZE=1>&copy; amdocs - Integrated Customer Management</FONT></TD></TR></TABLE></BODY></HTML>");

            //Close the HTML report
            fileStream.Flush();
            fileStream.Close();
            fileStream = null;
        }
        //*****************************************************************************************
        //*	End of function fnCloseTestSummary
        //*****************************************************************************************

        //*****************************************************************************************
        //*	Name		    : fnWriteToHtmlOutput
        //*	Description	    : The function Writes output to the HTML file
        //*	Author		    : Aniket Gadre
        //*	Input Params	: 	
        //*			            strDescription(String) - the description of the object
        //*			            strExpectedValue(String) - the expected value
        //*			            strObtainedValue(String) - the actual/obtained value
        //*			            strResult(String) - the result (Pass/Fail)
        //*	Return Values	: 	
        //*			            (Boolean) TRUE - Successful write
        //*				                  FALSE - Report file not created
        //*****************************************************************************************
        public void fnWriteToHtmlOutput(string strDescription, string strObtainedValue, string strResult)
        {

            //Declaring Variables
            string snapshotFilePath, sRowColor, strRelativePath;
            QualityCenter objQC = new QualityCenter();

            //Open the test case report for writing
            //Open the HTML file
            string g_strTestCaseReport = Environment.GetEnvironmentVariable("HTML_REPORTS_PATH") + "Report_" + g_strScriptName + ".html";
            StreamWriter fileStream = new StreamWriter(g_strTestCaseReport, true);

            //Increment the Operation Count
            g_OperationCount = g_OperationCount + 1;

            //Row Color
            if (g_OperationCount % 2 == 0)
            {
                sRowColor = "#EEEEEE";
            }
            else
            {
                sRowColor = "#D3D3D3";

            }

            //Check if the result is Pass or Fail
            if (strResult.ToUpper() == "PASS")
            {
                //Increment the Pass Count
                g_iPassCount++;

                //Increment the SnapShot count
                g_iSnapshotCount++;

                //Get the Full path of the snapshot

                snapshotFilePath = Environment.GetEnvironmentVariable("SCREEN_SHOT_PATH") + g_strScriptName + "\\SS_" + g_iSnapshotCount + ".jpg";
                strRelativePath = "Screen_Prints\\" + g_strScriptName + "\\SS_" + g_iSnapshotCount + ".jpg";

                //Capture the Snapshot
                fTakeScreenshot(snapshotFilePath);

                //Write the result into the file
                fileStream.WriteLine("<TR WIDTH=100%><TD  BGCOLOR=" + sRowColor + " WIDTH=10% ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B>" + g_OperationCount + "</B></FONT></TD><TD BGCOLOR=" + sRowColor + " WIDTH=30%><FONT FACE=VERDANA SIZE=2>" + strDescription + " </FONT></TD><TD BGCOLOR=" + sRowColor + " WIDTH=30%><FONT FACE=VERDANA SIZE=2>" + strObtainedValue + " </FONT></TD><TD BGCOLOR=" + sRowColor + " WIDTH=20% ALIGN=CENTER><A HREF='" + strRelativePath + "'><FONT FACE=VERDANA SIZE=2 COLOR=GREEN><B>" + strResult + " </B></FONT></A></TD></TR>");
                if (Global.objTD != null && Global.objTD.Connected == true)
                {
                    objQC.fQCStepUpdate("Step " + g_OperationCount.ToString(), strDescription, strObtainedValue, "PASSED");
                }
            }
            else
            {
                if (strResult.ToUpper() == "FAIL")
                {
                    //Increment the SnapShot count
                    g_iSnapshotCount++;

                    //Increment the Fail Count
                    g_iFailCount++;

                    //Get the Full path of the snapshot
                    snapshotFilePath = Environment.GetEnvironmentVariable("SCREEN_SHOT_PATH") + g_strScriptName + "\\SS_" + g_iSnapshotCount + ".jpg";
                    strRelativePath = "Screen_Prints\\" + g_strScriptName + "\\SS_" + g_iSnapshotCount + ".jpg";

                    //Capture the Snapshot
                    fTakeScreenshot(snapshotFilePath);

                    //Write the result into the file
                    fileStream.WriteLine("<TR WIDTH=100%><TD BGCOLOR=" + sRowColor + " WIDTH=10% ALIGN=CENTER><FONT FACE=VERDANA SIZE=2 ><B>" + g_OperationCount + "</B></FONT></TD><TD BGCOLOR=" + sRowColor + " WIDTH=30%><FONT FACE=VERDANA SIZE=2>" + strDescription + " </FONT></TD><TD BGCOLOR=" + sRowColor + " WIDTH=30%><FONT FACE=VERDANA SIZE=2>" + strObtainedValue + " </FONT></TD><TD BGCOLOR=" + sRowColor + " WIDTH=20% ALIGN=CENTER><A HREF='" + strRelativePath + "'><FONT FACE=VERDANA SIZE=2 COLOR=RED><B>" + strResult + " </B></FONT></A></TD></TR>");
                    if (Global.objTD != null && Global.objTD.Connected == true)
                    {
                        objQC.fQCStepUpdate("Step " + g_OperationCount.ToString(), strDescription, strObtainedValue, "FAILED");
                    }
                }
                else
                {
                    //Write Results into the file
                    fileStream.WriteLine("<TR WIDTH=100%><TD BGCOLOR=" + sRowColor + " WIDTH=10% ALIGN=CENTER><FONT FACE=VERDANA SIZE=2><B>" + g_OperationCount + "</B></FONT></TD><TD BGCOLOR=" + sRowColor + " WIDTH=30%><FONT FACE=VERDANA SIZE=2>" + strDescription + "</FONT></TD><TD BGCOLOR=" + sRowColor + " WIDTH=30%><FONT FACE=VERDANA SIZE=2>" + strObtainedValue + "</FONT></TD><TD BGCOLOR=" + sRowColor + " WIDTH=20% ALIGN=CENTER><FONT FACE=VERDANA SIZE=2 COLOR=orange><B>" + strResult + "</B></FONT></TD></TR>");
                    if (Global.objTD != null && Global.objTD.Connected == true)
                    {
                        objQC.fQCStepUpdate("Step " + g_OperationCount.ToString(), strDescription, strObtainedValue, "Done");
                    }
                }

            }


            //Close the HTML file
            //Close the HTML report
            fileStream.Flush();
            fileStream.Close();
            fileStream = null;


        }

        public void fTakeScreenshot(string SSPath)
        {
            using (Bitmap bmpScreenCapture = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height))
            {
                using (Graphics g = Graphics.FromImage(bmpScreenCapture))
                {
                    g.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, bmpScreenCapture.Size, CopyPixelOperation.SourceCopy);
                    bmpScreenCapture.Save(SSPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                }
            }
        }
    }
}
