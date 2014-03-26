using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using Scripting;
using System.Data.Odbc;
using System.Data;



namespace ASAP
{


    public class Driver
    {

        //Declare variables
        private Scripting.Dictionary GlobalDictionary = new Scripting.Dictionary();
        private Dictionary<string, string> OriginalDictionary = new Dictionary<string, string>();
        private Dictionary<int, string> Temp = new Dictionary<int, string>();
        private string strRunMode, strEnvCode, strVersion, strCalendarName, strEmailTo, strClearX;
        private string strAutoDP, strScreenShot, strHTMLReporting, strQCCommonPath, strFSCommonPath;
        private string strHTMLReportsPath, strDBSQLFromDataTable, strAutoScan, strTestPlanPath, strEmailFrom;
        private string strTestSetName = "";
        private string strTestSetPath = "";
        private string strReportsPath, strNewRun, strImagePath, strFSPath, strInfraVbs, strRecoveryScenarioPath;
        private string strCalendarsPath, strXLSWritePath, strExecutionLogPath, strStorage, strCalMainXlSPath, strDLLPath;
        private string strEXEPath, strINIPath, strVBSPath, strAPPVbs, strObjectReposPath, strReportingPath, strScreenShotPath;
        private string strTestSetIniFilePath, strEnvironmentXLSPath;
        private Dictionary<string, string> dict = new Dictionary<string, string>();
        private Reporting Reporter = new Reporting();
        private QualityCenter QC = new QualityCenter();
        private DBActivities objDB = new DBActivities();
        private Unix objUnix = new Unix();

        //***********************************List of Functions in Lib ******************************
        // 1.void fSetParams(string strRunModeTemp, string strEnvCodeTemp, string strVersionTemp, string strCalendarNameTemp, string strEmailToTemp, string strClearXTemp, string strAutoDPTemp, string strScreenShotTemp, string strHTMLReportingTemp, string strQCCommonPathTemp, string strFSCommonPathTemp, string strHTMLReportsPathTemp, string strDBSQLFromDataTableTemp, string strAutoScanTemp, string strTestPlanPathTemp, string strEmailFromTemp, string strReportsPathTemp, string strNewRunTemp)
        // 2.void fAppendPaths()
        // 3.bool fCopyEnvironmentXls()
        // 4.bool fReadReg(string strKeyPath)
        // 5.bool fWriteReg(string strKeyPath, string strKeyName, string strKeyValue)
        // 6.string fGetRunMode()
        // 7.void fTrimCalendarName()
        // 8.void fSetFilePaths()
        // 9.bool fHandleNewRunParam()
        //10.bool fCreateFilePaths()
        //11.void fAddInfoToEnvironment()
        //12.bool fCreateExecutionLogFilePath()
        //13.string fCreateTestSetIniFile()
        //14.bool fCopyCalendarAndCommonXls()
        //15.bool fDictionaryToScriptingDictionary(ref Scripting.Dictionary GD)
        //16.bool fScriptingDictionaryToDictionary(ref Scripting.Dictionary GD)
        //17.int fProcessDataFile(int rowID, ref string Skip)
        //18.bool fExecuteQueryToLoopCalendar(ref int iRows, ref string strScriptStartRows, ref string strTestCaseNames, ref string strScriptEndRows)
        //19.bool fGetReferenceData()
        //20.bool fSetReferenceData()
        //21.bool fUpdateTestCaseRowSkip(int row)
        //21a.bool fUpdateTestCaseRowSkip(int row, string strTestName, string strResult)
        //22.void fCreateHTMLSummaryReport()
        //23.void fCloseHTMLSummaryReport()
        //24.void fnCreateHtmlReport(string strFileName)
        //25.void fnCloseHtmlReport()
        //26.void fnWriteToHTMLOutput(string strDescription, string strObtainedValue, string strResult)
        //27.void fSetQCParams(TDAPIOLELib.TDConnection objTDConnect, string strTestSetNameTemp, string strTestSetPathTemp)
        //28.bool fAddTest(string strTestSetNameTemp, string strTestNameTemp)
        //29.bool fAttachResultsInQC(string strTestDetails)f
        //30.void fClearSkip(string sActionValue)
        //31.bool fDBActivities()
        //32.bool fBusPLinkExecuteUnixCommand()
        //33.Scripting.Dictionary fExecuteDBCheck()
        //34.bool fReallocationOfConsole()

        //***********************************List of Functions in Lib ******************************


        //*****************************************************************************************
        //*	Name		    : fSetParams
        //*	Description	    : Sets the params value
        //*	Author		    : Anil Agarwal
        //*	Input Params	: List of Parameters from the QTP script
        //*	Return Values	: None
        //*****************************************************************************************
        public void fSetParams(string strRunModeTemp, string strEnvCodeTemp, string strVersionTemp, string strCalendarNameTemp, string strEmailToTemp, string strClearXTemp, string strAutoDPTemp, string strScreenShotTemp, string strHTMLReportingTemp, string strQCCommonPathTemp, string strFSCommonPathTemp, string strHTMLReportsPathTemp, string strDBSQLFromDataTableTemp, string strAutoScanTemp, string strTestPlanPathTemp, string strEmailFromTemp, string strReportsPathTemp, string strNewRunTemp)
        {

            //Set the values in class variables
            strRunMode = strRunModeTemp;
            strEnvCode = strEnvCodeTemp;
            strVersion = strVersionTemp;
            strCalendarName = strCalendarNameTemp;
            strEmailTo = strEmailToTemp;
            strClearX = strClearXTemp;
            strAutoDP = strAutoDPTemp;
            strScreenShot = strScreenShotTemp;
            strHTMLReporting = strHTMLReportingTemp;
            strQCCommonPath = strQCCommonPathTemp;
            strFSCommonPath = strFSCommonPathTemp;
            strHTMLReportsPath = strHTMLReportsPathTemp;
            strDBSQLFromDataTable = strDBSQLFromDataTableTemp;
            strAutoScan = strAutoScanTemp;
            strTestPlanPath = strTestPlanPathTemp;
            strEmailFrom = strEmailFromTemp;
            strReportsPath = strReportsPathTemp;
            strNewRun = strNewRunTemp;
        }

        //*****************************************************************************************
        //*	Name		    : fAppendPaths
        //*	Description	    : Sets the paths
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: Bool True on Success / False on failure
        //*****************************************************************************************
        public void fAppendPaths()
        {
            //Add a "/" to the path in the end to QC common path if not present
            if (!strQCCommonPath.Equals(""))
            {
                if (strQCCommonPath.Substring(strQCCommonPath.Length - 1).Equals("\\") == false)
                {
                    strQCCommonPath = strQCCommonPath + "\\";
                }
            }


            //Add a "/" to the path in the end to FS common path if not present
            if (strFSCommonPath.Substring(strFSCommonPath.Length - 1).Equals("\\") == false)
            {
                strFSCommonPath = strFSCommonPath + "\\";
            }

            //Add a "/" to the path in the end to Reports path if not present
            if (!strReportsPath.Equals(""))
            {
                if (strReportsPath.Substring(strReportsPath.Length - 1).Equals("\\") == false)
                {
                    strReportsPath = strReportsPath + "\\";
                }
            }

        }

        //*****************************************************************************************
        //*	Name		    : fCopyEnvironmentXls
        //*	Description	    : Copies the environment xls and stores in a temp folder
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: Bool True on Success / False on failure
        //*****************************************************************************************
        public bool fCopyEnvironmentXls()
        {
            try
            {
                //Check of the environment file exists 
                if (!System.IO.File.Exists(strFSCommonPath + "Environments.xls"))
                {
                    //return false
                    return false;
                }

                //If Temp folder is not there, create the same
                if (!Directory.Exists(strFSCommonPath + "Temp\\"))
                {
                    Directory.CreateDirectory(strFSCommonPath + "Temp\\");
                }

                //Set the path for backup Environment.xls file
                string strTmpEnvPath = strFSCommonPath + "Temp\\" + DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + "_tempEnvironments.xls";

                //Copy the environment file to the backup environment file
                System.IO.File.Copy(strFSCommonPath + "Environments.xls", strTmpEnvPath);

                //return success
                return true;
            }
            catch (Exception e)
            {
                return false;
            }

        }


        //*****************************************************************************************
        //*	Name		    : fReadReg
        //*	Description	    : Reads the Registry and checks if a key is present or not
        //*	Author		    : Anil Agarwal
        //*	Input Params	: String strKeyName
        //*	Return Values	: Bool True on Success / False on failure
        //*****************************************************************************************
        public bool fReadReg(string strKeyPath)
        {
            try
            {

                Global.fUpdateExecutionLog("*********************************************************************************************");
                Global.fUpdateExecutionLog("Execution Time is : " + DateTime.Now);
                Global.fUpdateExecutionLog("Tester Name is : " + Environment.UserName);
                Global.fUpdateExecutionLog("Machine Name is : " + Environment.MachineName);
                Global.fUpdateExecutionLog("*********************************************************************************************");

                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fReadReg");
                Microsoft.Win32.RegistryKey key;
                key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(strKeyPath);

                //if key is null that means key does not exist
                if (key == null)
                {
                    //Write the code to enter the value in the Execution log
                    Global.fUpdateExecutionLog(LogType.debug, "GlobalDictionary Entry not found in Registry");
                    //return failure
                    return false;
                }
                else
                {
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog(LogType.info, "GlobalDictionary Entry already Present in Registry");
                    //return success
                    return true;
                }
            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while checking GlobalDictionary Entry in Registry. Exception is " + e);
                return false;
            }

        }

        //*****************************************************************************************
        //*	Name		    : fWriteReg
        //*	Description	    : Writes the key in Registry
        //*	Author		    : Anil Agarwal
        //*	Input Params	: String strKeyName, String strKeyValue, String strKeyType
        //*	Return Values	: Bool True on Success / False on failure
        //*****************************************************************************************
        public bool fWriteReg(string strKeyPath, string strKeyName, string strKeyValue)
        {
            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fWriteReg");

                Microsoft.Win32.RegistryKey rKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(strKeyPath, Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree);
                //Microsoft.Win32.Registry.CurrentUser.OpenSubKey(strKeyName, true);
                rKey.SetValue(strKeyName, strKeyValue);

                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.info, "Successfully created key " + strKeyName + " with Key value " + strKeyValue + " in Registry");

                //return success
                return true;
            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while writing GlobalDictionary Entry in Registry. Exception is " + e);
                return false;
            }

        }

        //*****************************************************************************************
        //*	Name		    : fGetRunMode
        //*	Description	    : Get the Run Mode
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: String RunMode
        //*****************************************************************************************
        public string fGetRunMode()
        {
            //Update the Execution log
            Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fGetRunMode");

            //Based on the Run mode 
            switch (strRunMode.ToUpper())
            {
                //Based on the strRunMode return the run mode value
                case "DEV":
                    {
                        strRunMode = "DEVELOPMENT";
                        return strRunMode;
                    }

                case "PROD":
                    {
                        strRunMode = "PRODUCTION";
                        return strRunMode;
                    }

                case "TRAINING":
                    {
                        strRunMode = "TRAINING";
                        return strRunMode;
                    }

                default:
                    {
                        strRunMode = "DEVELOPMENT";
                        return strRunMode;
                    }

            }
        }

        //*****************************************************************************************
        //*	Name		    : fTrimCalendarName
        //*	Description	    : Trim the calendar name and remove the extension if present
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: String strCalendarName
        //*****************************************************************************************
        public void fTrimCalendarName()
        {
            //if the calendar name contains .xls then remove the same
            if (strCalendarName.ToLower().Contains(".xls"))
            {
                strCalendarName = strCalendarName.Substring(0, (strCalendarName.Length - 4));
            }
            else if (strCalendarName.ToLower().Contains(".xlsx"))
            {
                strCalendarName = strCalendarName.Substring(0, (strCalendarName.Length - 5));
            }

        }

        //*****************************************************************************************
        //*	Name		    : fSetFilePaths
        //*	Description	    : Sets all file paths
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: String strCalendarName
        //*****************************************************************************************
        public void fSetFilePaths()
        {
            Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fSetFilePaths");

            //Set QC Test Plan Path
            strTestPlanPath = strQCCommonPath + "Automation_Coverage\\";

            //if version is populated then Set Image Path and FSCommon path
            if (!strVersion.Equals(""))
            {
                //Set Image and File Ss
                strImagePath = strQCCommonPath + strVersion + "\\Screenshot\\";
                //strFSPath = strFSCommonPath + strVersion + "\\";
            }
            else
            {
                strImagePath = strQCCommonPath + strEnvCode + "\\Screenshot\\";
                //strFSPath = strFSCommonPath;
            }

            //Set Recovery and Infra VBS Paths
            strInfraVbs = strFSCommonPath + "General Activities\\Infra\\";
            strRecoveryScenarioPath = strFSCommonPath + "General Activities\\Recovery Scenario\\";

            //Set the calendars path
            strCalendarsPath = strFSPath + "DATA_FILES_PER_CALENDAR\\";

            //Write code for xls write path if QC is connected later
            if (Global.objTD != null && Global.objTD.Connected == true)
            {
                strXLSWritePath = strCalendarsPath + strRunMode + "\\" + strEnvCode + "\\" + strTestSetName + "\\" + strCalendarName + "\\";
            }
            else
            {
                strXLSWritePath = strCalendarsPath + strRunMode + "\\" + strEnvCode + "\\" + strCalendarName + "\\";
            }

            //Other paths
            strStorage = strFSPath + "STORAGE\\";
            strCalMainXlSPath = strStorage + "CALENDARS_MAIN_EXCEL\\" + strRunMode + "\\";
            strDLLPath = strStorage + "DLL\\";
            strEXEPath = strStorage + "EXE\\";
            strINIPath = strStorage + "INI\\";
            strVBSPath = strStorage + "Libraries\\";
            strAPPVbs = strVBSPath + "App\\";
            strObjectReposPath = strStorage + "Object_Repository\\";
            strEnvironmentXLSPath = strFSCommonPath + "Environments.xls";

            //Set the strReportingPath
            if (!strReportsPath.Equals(""))
            {
                strReportingPath = strReportsPath + "Automation_Results\\";
            }
            else
            {
                strReportingPath = strCalendarsPath + "Automation_Results\\";
            }


            //Later on add the code when  QC is connected
            if (Global.objTD != null && Global.objTD.Connected == true)
            {
                strReportingPath = strReportingPath + strRunMode + "\\" + strEnvCode + "\\" + strTestSetName + "\\" + strCalendarName + "\\";
            }
            else
            {
                strReportingPath = strReportingPath + strRunMode + "\\" + strEnvCode + "\\" + strCalendarName + "\\";
            }

            //Get the current date and Current TimeStamp
            string strCurrentDate = DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year;
            string strCurrentTimeStamp = "Time_" + DateTime.Now.Hour + "-" + DateTime.Now.Minute + "-" + DateTime.Now.Second;

            //Append Date TIme to Reports path
            strReportingPath = strReportingPath + strCurrentDate + "\\" + strCurrentTimeStamp + "\\";

            //If the HTMLReports flag is Y then set the path for HTMLReports
            if (strHTMLReporting.ToUpper().Equals("Y"))
            {
                strHTMLReportsPath = strReportingPath + "HTML_REPORTS\\";
            }
            else
            {
                strHTMLReportsPath = "";
            }

            //if the Screenshot path flag is Y then set the path for Screenshots path
            if (strScreenShot.ToUpper().Equals("Y"))
            {
                strScreenShotPath = strHTMLReportsPath + "SCREEN_PRINTS\\";
            }
            else
            {
                strScreenShotPath = "";
            }

        }

        //*****************************************************************************************
        //*	Name		    : fHandleNewRunParam
        //*	Description	    : Handles the new run param and delete execution folder is set to Y
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: bool True if success/false if fail
        //*****************************************************************************************
        public bool fHandleNewRunParam()
        {
            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fHandleNewRunParam");

                //if the new run param is Y then delete the execution folder
                if (strNewRun.ToUpper().Equals("Y"))
                {
                    //delete the folder and all its content
                    if (Directory.Exists(strXLSWritePath))
                    {
                        //Delete all its content
                        Directory.Delete(strXLSWritePath);
                    }

                }

                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.info, "Successfully Deleted the Execution Folder");

                //return success
                return true;

            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while deleting the Execution Folder. Exception is " + e);
                return false;
            }
        }

        //*****************************************************************************************
        //*	Name		    : fCreateFilePaths
        //*	Description	    : Create the file paths
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: bool True if success/false if fail
        //*****************************************************************************************
        public bool fCreateFilePaths()
        {
            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fCreateFilePaths");

                string[] arrFolders = { strXLSWritePath, strHTMLReportsPath, strScreenShotPath };

                //Loop through the array and create the folder structure
                for (int i = 0; i < arrFolders.Length; i++)
                {
                    Global.fUpdateExecutionLog(LogType.info, "Creating path: " + arrFolders[i]);
                    //Create the paths if not null
                    if (!arrFolders[i].Equals("") && !Directory.Exists(arrFolders[i]))
                    {
                        Directory.CreateDirectory(arrFolders[i]);
                    }

                }

                //return success
                return true;

            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while Creating the folder structure. Exception is " + e);
                return false;
            }
        }

        //*****************************************************************************************
        //*	Name		    : fAddInfoToEnvironment
        //*	Description	    : Add Info to Environment
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: bool True if success/false if fail
        //*****************************************************************************************
        public void fAddInfoToEnvironment()
        {
            //Update the Execution log
            Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fAddInfoToEnvironment");

            //Add all the parameter value to the Dictionary object
            Environment.SetEnvironmentVariable("RUN_MODE", strRunMode);
            Environment.SetEnvironmentVariable("VERSION", strVersion);
            Environment.SetEnvironmentVariable("ENV_KEY", strEnvCode);
            Environment.SetEnvironmentVariable("CALENDAR_NAME", strCalendarName);
            Environment.SetEnvironmentVariable("TEST_SET_INI_NAME", "TestSet.ini");
            Environment.SetEnvironmentVariable("IMAGE_PATH", strImagePath);
            Environment.SetEnvironmentVariable("TEST_PLAN_PATH", strTestPlanPath);
            Environment.SetEnvironmentVariable("ENV_XLS", strEnvironmentXLSPath);
            Environment.SetEnvironmentVariable("INFRA_PATH", strInfraVbs);
            Environment.SetEnvironmentVariable("APP_VBS_PATH", strAPPVbs);
            Environment.SetEnvironmentVariable("RECOVERY_SCENARIO_PATH", strRecoveryScenarioPath);
            Environment.SetEnvironmentVariable("OBJECT_REPOSITORY_PATH", strObjectReposPath);
            Environment.SetEnvironmentVariable("STORAGE_PATH", strStorage);
            Environment.SetEnvironmentVariable("DB_SQL_FROM_DATATABLE", strDBSQLFromDataTable);
            Environment.SetEnvironmentVariable("AUTO_DP", strAutoDP);
            Environment.SetEnvironmentVariable("DLL_PATH", strDLLPath);
            Environment.SetEnvironmentVariable("EXE_PATH", strEXEPath);
            Environment.SetEnvironmentVariable("INI_PATH", strINIPath);
            Environment.SetEnvironmentVariable("CALENDARS_MAIN_EXCEL_PATH", strCalMainXlSPath);
            Environment.SetEnvironmentVariable("RO_COMMON_XLS", strCalMainXlSPath + "COMMON.xls");
            Environment.SetEnvironmentVariable("RO_MAIN_XLS", strCalMainXlSPath + strCalendarName + ".xls");
            Environment.SetEnvironmentVariable("CALENDARS_PATH", strCalendarsPath);
            Environment.SetEnvironmentVariable("XLS_WRITE_PATH", strXLSWritePath);
            Environment.SetEnvironmentVariable("TEST_SET_INI", strXLSWritePath + "TestSet.ini");
            Environment.SetEnvironmentVariable("W_COMMON_XLS", strXLSWritePath + "COMMON.xls");
            Environment.SetEnvironmentVariable("W_MAIN_XLS", strXLSWritePath + strCalendarName + ".xls");
            Environment.SetEnvironmentVariable("HTML_REPORTS_PATH", strHTMLReportsPath);
            Environment.SetEnvironmentVariable("SCREEN_SHOT_PATH", strScreenShotPath);
            Environment.SetEnvironmentVariable("MAIN_SHEET", "[MAIN$]");
            Environment.SetEnvironmentVariable("KEEP_REFER_SHEET", "[KEEP_REFER$]");
            Environment.SetEnvironmentVariable("DB_SQL_SHEET", "[DB_SQL$]");
            Environment.SetEnvironmentVariable("UNIX_SHEET", "[UNIX$]");
            Environment.SetEnvironmentVariable("JOBS_SHEET", "[JOBS$]");
            Environment.SetEnvironmentVariable("TEST_SET_NAME", strTestSetName);
            Environment.SetEnvironmentVariable("TEST_SET_PATH", strTestSetPath);


            //Update the Execution log
            Global.fUpdateExecutionLog(LogType.info, "Added Info to the Environment Object");

        }

        //*****************************************************************************************
        //*	Name		    : fCreateExecutionLogFilePath
        //*	Description	    : Create the Execution Log File path
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: bool True if success/false if fail
        //*****************************************************************************************
        public bool fCreateExecutionLogFilePath()
        {
            try
            {
                //if version is populated then Set Image Path and FSCommon path
                if (!strVersion.Equals(""))
                {
                    //Set FSPath
                    strFSPath = strFSCommonPath + strVersion + "\\";
                }
                else
                {
                    strFSPath = strFSCommonPath;
                }

                strExecutionLogPath = strFSPath + "EXECUTION_LOGS\\";

                //Create the Execution Log path if not present
                if (!Directory.Exists(strExecutionLogPath))
                {
                    //Create the execution Log path
                    Directory.CreateDirectory(strExecutionLogPath);
                }


                //Create the Execution Log file path
                strExecutionLogPath = strExecutionLogPath + "Daily_Log_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + ".txt";
                StreamWriter streamWriter;
                //If the file does not exist create the same
                if (!System.IO.File.Exists(strExecutionLogPath))
                {
                    //Create the file
                    streamWriter = new StreamWriter(strExecutionLogPath);

                    using (streamWriter)
                    {
                        //Write the header information
                        streamWriter.WriteLine("********************************************************************************************");
                        streamWriter.WriteLine("*************************************** Execution Logs **************************************");
                        streamWriter.WriteLine("*********************************************************************************************");

                    }

                }

                Environment.SetEnvironmentVariable("EXECUTION_LOG_FILE_PATH", strExecutionLogPath);

                //Return success
                return true;
            }
            catch (Exception e)
            {
                return false;
            }

        }

        //*****************************************************************************************
        //*	Name		    : fCreateTestSetIniFile
        //*	Description	    : Create the Execution Log File path
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: bool True if success/false if fail
        //*****************************************************************************************
        public string fCreateTestSetIniFile()
        {
            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fCreateTestSetIniFile");

                DBActivities objDB = new DBActivities();
                OdbcConnection objConn = new OdbcConnection();
                objConn = objDB.fConnectToXLS(strEnvironmentXLSPath);

                //Handle null if connection fails
                if (objConn == null)
                {
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog(LogType.error, "DB Connection to environment xls failed. Path: " + strEnvironmentXLSPath);
                    return "";
                }

                //Open dataset
                DataSet objDS = new DataSet();
                string SQL = "Select * from [ENVIRONMENTS$] Where Environment = '" + strEnvCode + "'";
                objDS = objDB.fExecuteSelectQuery(SQL, objConn);

                //handle null if exeuction of query fails
                if (objDS == null)
                {
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog(LogType.error, "Fetching details for Env: " + strEnvCode + " failed");
                    objConn.Close();
                    return "";
                }

                //Check if any rows returned
                DataTable objDT = objDS.Tables[0];
                if (objDT.Rows.Count == 0)
                {
                    //Log
                    Global.fUpdateExecutionLog(LogType.error, "No environment details fetched by query " + SQL);
                    objConn.Close();
                    return "";
                }

                //Loop through dataSet and fetch the values in env var
                int iFieldsCnt = objDS.Tables[0].Columns.Count;

                for (int i = 0; i < iFieldsCnt; i++)
                {
                    Environment.SetEnvironmentVariable(objDS.Tables[0].Columns[i].ColumnName.ToUpper().Trim(), objDS.Tables[0].Rows[0].ItemArray[i].ToString());
                }

                //Close datasets and connections
                objDS.Dispose();
                objConn.Close();


                //Create the testSet ini file path
                strTestSetIniFilePath = strXLSWritePath + "TestSet.ini";
                StreamWriter streamWriter;
                //If the file does not exist create the same
                if (System.IO.File.Exists(strTestSetIniFilePath))
                {
                    //Delete the TestSet ini file
                    System.IO.File.Delete(strTestSetIniFilePath);
                }

                //Create the file
                streamWriter = new StreamWriter(strTestSetIniFilePath);

                //Write the header information
                using (streamWriter)
                {
                    streamWriter.WriteLine("[Environment]");
                }

                //Getting all the keys 
                System.Collections.IDictionary keyTemp = Environment.GetEnvironmentVariables();

                
                //Loop though the Dictionary object and enter all its details in the testSet.ini file
                foreach (System.Collections.DictionaryEntry de in keyTemp)
                {
                    //Open the file in append mode
                    //Write the key and value in TestSet.ini file
                    using (streamWriter = new StreamWriter(strTestSetIniFilePath, true))
                    {
                        streamWriter.WriteLine(de.Key + "=" + de.Value);
                    }
                }
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.info, "TestSet.ini file created successfully. Path is " + strTestSetIniFilePath);
                //Return success
                return strTestSetIniFilePath;
            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while creating the TestSet.ini file. Exception is " + e);
                return e.StackTrace;
            }
        }

        //*****************************************************************************************
        //*	Name		    : fCopyCalendarAndCommonXls
        //*	Description	    : Copies the Calendar and Common Xls to the Execution Folder
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: bool True if success/false if fail
        //*****************************************************************************************
        public bool fCopyCalendarAndCommonXls()
        {
            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fCopyCalendarAndCommonXls");

                //Check if the Calendar xls exists in execution folder
                if (!System.IO.File.Exists(Environment.GetEnvironmentVariable("W_MAIN_XLS")))
                {
                    //Check if the Calendar is present in Storage
                    if (!System.IO.File.Exists(Environment.GetEnvironmentVariable("RO_MAIN_XLS")))
                    {
                        //Reporting failure

                        return false;
                    }

                    //Copy the file from storage to execution folder
                    System.IO.File.Copy(Environment.GetEnvironmentVariable("RO_MAIN_XLS"), Environment.GetEnvironmentVariable("W_MAIN_XLS"));
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog("Calendar File Copied successfully to the execution folder");

                }


                //if the DB_SQL_FROM_DATATABLE is Y then only Copy the Common xls
                if (!strDBSQLFromDataTable.ToUpper().Equals("Y"))
                {
                    //Check if the Common xls exists in execution folder
                    if (!System.IO.File.Exists(Environment.GetEnvironmentVariable("W_COMMON_XLS")))
                    {
                        //Check if the Calendar is present in Storage
                        if (!System.IO.File.Exists(Environment.GetEnvironmentVariable("RO_COMMON_XLS")))
                        {
                            //Reporting failure

                            return false;
                        }

                        //Copy the file from storage to execution folder
                        System.IO.File.Copy(Environment.GetEnvironmentVariable("RO_COMMON_XLS"), Environment.GetEnvironmentVariable("W_COMMON_XLS"));
                        //call the function to update Execution Log
                        Global.fUpdateExecutionLog(LogType.info, "Common xls file Copied successfully to the execution folder");

                    }
                }

                //return success
                return true;

            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while Copying the Calendar and Common xls. Exception is " + e);

                //return false
                return false;
            }

        }

        //*****************************************************************************************
        //*	Name		    : fDictionaryToScriptingDictionary
        //*	Description	    : Converts the dictionary object to scripting.dictionary object
        //*	Author		    : Anil Agarwal
        //*	Input Params	: Dictionary object
        //*	Return Values	: Scripting.Dictionary if success/null if fail
        //*****************************************************************************************
        public bool fDictionaryToScriptingDictionary(ref Scripting.Dictionary GD)
        {
            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fDictionaryToScriptingDictionary");

                //Remove all the objects from Dictionary object
                GD.RemoveAll();

                //Loop through the dictionary object and add it to the Scripting.Dictionary
                foreach (KeyValuePair<string, string> entry in Global.Dictionary)
                {
                    //Add the entry in GlobalDictionary object of type Scripting.Dictionary
                    GD.Add(entry.Key, entry.Value);
                }

                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.info, "Dictionary object successfully converted to Scripting.Dictionary object");

                //return true if successful
                return true;

            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while converting the Dictionary object to Scripting.Dictionary object. Exception is " + e);
                return false;
            }

        }

        //*****************************************************************************************
        //*	Name		    : fDictionaryToScriptingDictionary
        //*	Description	    : Converts the dictionary object to scripting.dictionary object
        //*	Author		    : Anil Agarwal
        //*	Input Params	: Dictionary object
        //*	Return Values	: Scripting.Dictionary if success/null if fail
        //*****************************************************************************************
        public bool fDictionaryToScriptingDictionary(ref Scripting.Dictionary GD, Dictionary<string, string> dictTemp)
        {
            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fDictionaryToScriptingDictionary");

                //Remove all the objects from Dictionary object
                GD.RemoveAll();

                //Loop through the dictionary object and add it to the Scripting.Dictionary
                foreach (KeyValuePair<string, string> entry in dictTemp)
                {
                    //Add the entry in GlobalDictionary object of type Scripting.Dictionary
                    GD.Add(entry.Key, entry.Value);
                }

                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.info, "Dictionary object successfully converted to Scripting.Dictionary object");

                //return true if successful
                return true;

            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while converting the Dictionary object to Scripting.Dictionary object. Exception is " + e);
                return false;
            }

        }

        //*****************************************************************************************
        //*	Name		    : fScriptingDictionaryToDictionary
        //*	Description	    : Converts the dictionary object to scripting.dictionary object
        //*	Author		    : Anil Agarwal
        //*	Input Params	: Dictionary object
        //*	Return Values	: Scripting.Dictionary if success/null if fail
        //*****************************************************************************************
        public bool fScriptingDictionaryToDictionary(ref Scripting.Dictionary GD)
        {
            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fScriptingDictionaryToDictioanry");

                //Remove all the objects from Dictionary object
                //Global.Dictionary.Clear();

                //Define a variable
                string strKeys = "";
                //Loop thorugh all the keys in the Scripting.Dictionary object

                Object[] gdKeys = (Object[])GD.Keys();

                foreach (Object Key in gdKeys)
                {
                    
                    //Global.fUpdateExecutionLog(LogType.debug, "Key : " + Key.ToString());
                    strKeys = strKeys + "^" + Key.ToString();
                }

                //Remove the ";" in the beginning
                strKeys = strKeys.Substring(1);

                //Define a variable
                string strItems = "";

                Object[] gdItems = (Object[])GD.Items();


                //Loop thorugh all the keys in the Scripting.Dictionary object
                foreach (Object Items in gdItems)
                {
                    if (Items == null)
                    {
                       // Global.fUpdateExecutionLog(LogType.debug, "Item: is null");
                        strItems = strItems + "^" + "";
                    }
                    else
                    {
                        //Global.fUpdateExecutionLog(LogType.debug, "Item: " + Items.ToString());
                        strItems = strItems + "^" + Items.ToString();
                    }
                }

                //Remove the ";" in the beginning
                strItems = strItems.Substring(1);

                //Split on the basis of ";" and get the keys and items in an array
                string[] arrKeys = strKeys.Split('^');
                string[] arrItems = strItems.Split('^');

                //Loop and add the values in the Dictionary object
                for (int k = 0; k < arrKeys.Length; k++)
                {
                    if (Global.Dictionary.ContainsKey(arrKeys[k])) Global.Dictionary[arrKeys[k]] = arrItems[k];
                    else Global.Dictionary.Add(arrKeys[k], arrItems[k]);

                    //Global.fUpdateExecutionLog(LogType.info, "Added Key: " + arrKeys[k] + " and value: " + arrItems[k]);
                }


                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.info, "Succesfully converted the Scripting.Dictionary object to Dictionary object");

                //return in case of success
                return true;

            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while converting the Scripting.Dictionary object to Dictionary object. Exception is " + e);
                return false;
            }

        }

        //***********************************************************************
        // Name: 	     fProcessDataFile
        // Description:  Reads the Excel and gets all values in Global Hash
        // Author:       Aniket Gadre
        // Input Params: conn - DB Connection Object
        //               rowID - row number which is to be processed
        // Return Value: 1 or 0 depending on the row type
        //***********************************************************************
        public int fProcessDataFile(int rowID, ref string Skip)
        {
            //string key;
            DBActivities objDB = new DBActivities();
            OdbcConnection objConn = new OdbcConnection();

            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fProcessDataFile");

                //Establish connection to the Calendar xls
                objConn = objDB.fConnectToXLS(Environment.GetEnvironmentVariable("W_MAIN_XLS"));

                //Handle null if connection fails
                if (objConn == null)
                {
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog(LogType.error, "Failed to establish connection to the calendar file. Excel path is " + Environment.GetEnvironmentVariable("W_MAIN_XLS"));

                    return -1;
                }

                int iret = 0;

                //Query the excel
                DataSet objDS = new DataSet();
                string SQL = "Select * from [MAIN$] where ID = '" + rowID + "'";
                objDS = objDB.fExecuteSelectQuery(SQL, objConn);

                //handle null
                if (objDS == null)
                {
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog(LogType.error, "Processing Data Sheet failed. Null Dataset returned");
                    objConn.Close();
                    return -1;
                }

                //Set data row
                DataRow record = objDS.Tables[0].Rows[0];
                int iCounter = 1;

                if (record["HEADER_IND"] != null && record["HEADER_IND"].ToString() == "HEADER")
                {
                    //Delete already present keys in GD
                    Global.Dictionary.Clear();

                    //Delete already present keys in temp 
                    Temp.Clear();

                    //Looping through each field in record
                    foreach (DataColumn col in objDS.Tables[0].Columns)
                    {
                        if (record[col].ToString() != "")
                        {
                            Temp.Add(iCounter, record[col].ToString());
                            Global.Dictionary.Add(record[col].ToString(), "");
                            iret = 0;
                            iCounter++;
                        }
                    }
                }
                else
                {
                    //Set the item for each key in the GlobalDictionary to null
                    //foreach (string key in Global.Dictionary.Keys) Global.Dictionary[key] = "";

                    //Looping through each field in record
                    foreach (DataColumn col in objDS.Tables[0].Columns)
                    {
                        if (Temp.ContainsKey(iCounter - 1))
                        {
                            Global.Dictionary[Temp[iCounter - 1]] = record[col].ToString();
                        }

                        iret = 1;
                        iCounter++;
                    }

                }

                //Set skip value
                Skip = record["SKIP"].ToString();
                objDS.Dispose();
                objConn.Close();
                objDB = null;

                //return
                return iret;

            }
            catch (Exception e)
            {
                //Close the connection
                objConn.Close();
                Global.fUpdateExecutionLog(LogType.error, "Got exception while executing the fProcessDataFile function. Exception is " + e);
                return -1;
            }


        }


        //***********************************************************************
        // Name: 	     fExecuteQueryToLoopCalendar
        // Description:  Executes the query that Loops the calendar
        // Author:       Aniket Gadre
        // Input Params: None
        // Out Params:   int Row- No.of Records, int iStartRow - The start row
        //               int iEndRow - The end row number
        // Return Value: None
        //***********************************************************************
        public bool fExecuteQueryToLoopCalendar(ref int iRows, ref string strScriptStartRows, ref string strTestCaseNames, ref string strScriptEndRows)
        {
            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fExecuteQueryToLoopCalendar");

                DBActivities objDB = new DBActivities();
                OdbcConnection objConn = new OdbcConnection();
                DataSet objDS = new DataSet();


                //Establish connection to the Calendar xls
                objConn = objDB.fConnectToXLS(Environment.GetEnvironmentVariable("W_MAIN_XLS"));

                //Handle null if connection fails
                if (objConn == null)
                {
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog(LogType.error, "Failed to establish connection to the calendar file. Excel path is " + Environment.GetEnvironmentVariable("W_MAIN_XLS"));
                    return false;
                }

                //Declare variables
                string SQL;

                //SQL 
                SQL = "SELECT int(B.ID) As [iCurHID], TEST_NAME, min(A.Nextal) As [iNextHID] FROM [MAIN$] as B, (SELECT int(ID) As [Nextal] FROM [MAIN$] WHERE HEADER_IND Is Not Null GROUP BY ID, TEST_NAME) as A WHERE A.Nextal > int(B.ID) And B.HEADER_IND Is Not Null and SKIP is null group by int(B.ID), TEST_NAME order by int(B.ID)";
                objDS = objDB.fExecuteSelectQuery(SQL, objConn);

                //handle null if query fails
                if (objDS == null)
                {
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog(LogType.error, "Processing Data Sheet failed. Null Dataset returned");
                    objConn.Close();

                    return false;
                }

                //Get Table row count
                iRows = objDS.Tables[0].Rows.Count;

                //If no records found, exit the loop
                if (iRows == 0)
                {
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog(LogType.info, "No Records found to execute in the calendar");
                    objConn.Close();
                    return false;
                }

                String strScriptStartRowsTemp, strTestCaseNamesTemp, strScriptEndRowsTemp;

                //Loop though all the rows to and store their values in respective varaibles
                for (int j = 0; j < iRows; j++)
                {
                    //Concateante script start rows
                    strScriptStartRowsTemp = objDS.Tables[0].Rows[j].ItemArray[0].ToString();
                    strScriptStartRows = strScriptStartRows + ";" + strScriptStartRowsTemp;

                    //Concatenate Script end rows
                    strScriptEndRowsTemp = objDS.Tables[0].Rows[j].ItemArray[2].ToString();
                    strScriptEndRows = strScriptEndRows + ";" + strScriptEndRowsTemp;

                    //Concatename Test Case names
                    strTestCaseNamesTemp = objDS.Tables[0].Rows[j].ItemArray[1].ToString();
                    strTestCaseNames = strTestCaseNames + ";" + strTestCaseNamesTemp;
                }

                //Remove the extra ";" in the beginning
                strScriptStartRows = strScriptStartRows.Substring(1);
                strScriptEndRows = strScriptEndRows.Substring(1);
                strTestCaseNames = strTestCaseNames.Substring(1);

                //Close connections
                objDS.Dispose();
                objConn.Close();

                return true;

            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while executing the function fExecuteQueryToLoopCalendar. Exception is " + e);

                return false;
            }
        }

        //***********************************************************************
        // Name: 	      fGetReferenceData
        // Description:   Replaces Reference data from KEEP_REFER Sheet
        // Author:        Aniket Gadre
        // Input Params	: DB connection objecyt
        // Return Values: None
        //***********************************************************************
        public bool fGetReferenceData()
        {

            DBActivities objDB = new DBActivities();
            OdbcConnection objConn = new OdbcConnection();
            DataSet objDS = new DataSet();

            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fGetReferenceData");

                //Establish connection to the Calendar xls
                objConn = objDB.fConnectToXLS(Environment.GetEnvironmentVariable("W_MAIN_XLS"));

                //Handle null if connection fails
                if (objConn == null)
                {
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog(LogType.error, "Failed to establish connection to the calendar file. Excel path is " + Environment.GetEnvironmentVariable("W_MAIN_XLS"));

                    return false;
                }

                string SQL = "";
                string Key_Name = "";
                char[] delimiter = { '&' };

                //Connect to Excel sheet
                Console.WriteLine("Getting Reference Data");

                //Fetch keys and values of GD
                int iCnt = Global.Dictionary.Count;

                //Clear Original Dictionary
                OriginalDictionary.Clear();

                //Copying dict in temp
                foreach (KeyValuePair<string, string> KVP in Global.Dictionary) OriginalDictionary.Add(KVP.Key, KVP.Value);

                //Key Collection
                Dictionary<string, string>.KeyCollection keys = OriginalDictionary.Keys;

                string startChar = "";

                //Loop through Keys
                foreach (string key in keys)
                {
                    startChar = "";

                    //Check 1st character of value
                    if (OriginalDictionary[key] != "") startChar = OriginalDictionary[key].Substring(0, 1);

                    //if 1st char == &
                    if (startChar == "&")
                    {
                        //Key Name
                        Key_Name = OriginalDictionary[key].Split(delimiter)[1];

                        //SQL
                        SQL = "Select KEY_VALUE FROM [KEEP_REFER$] WHERE KEY_NAME = '" + Key_Name + "'";

                        //Execute Query
                        objDS = objDB.fExecuteSelectQuery(SQL, objConn);

                        //handle null
                        if (objDS == null) continue;

                        //CHeck rows
                        if (objDS.Tables[0].Rows.Count == 0)
                        {
                            Global.fUpdateExecutionLog(LogType.error, "No rows returned by query " + SQL);
                            objConn.Close();
                            return false;
                        }

                        //Replace Value
                        if (objDS.Tables[0].Rows[0]["KEY_VALUE"] == null) Global.Dictionary[key] = "";
                        else Global.Dictionary[key] = objDS.Tables[0].Rows[0]["KEY_VALUE"].ToString();
                    }
                }
                objDS.Dispose();
                objConn.Close();

                //return success
                return true;
            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while executing fGetReferenceData function. Exception is " + e);
                objConn.Close();
                return false;
            }
        }

        //***********************************************************************
        // Name: 	      fSetReferenceData
        // Description:   Sets Reference data in KEEP_REFER Sheet
        // Author:        Aniket Gadre
        // Input Params	: DB Connection Object
        // Return Values: None
        //***********************************************************************
        public bool fSetReferenceData()
        {

            DBActivities objDB = new DBActivities();
            OdbcConnection objConn = new OdbcConnection();
            DataSet objDS = new DataSet();

            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fSetReferenceData");

                //Establish connection to the Calendar xls
                objConn = objDB.fConnectToXLS(Environment.GetEnvironmentVariable("W_MAIN_XLS"));

                //Handle null if connection fails
                if (objConn == null)
                {
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog(LogType.error, "Failed to establish connection to the calendar file. Excel path is " + Environment.GetEnvironmentVariable("W_MAIN_XLS"));

                    return false;
                }
                string SQL, sSQL;
                string Key_Name, Key_Value;
                char[] delimiter = { '@' };

                //Connect to Excel sheet

                //Key Collection
                Dictionary<string, string>.KeyCollection keys = OriginalDictionary.Keys;

                string startChar = "";

                //Loop through Keys
                foreach (string key in keys)
                {
                    startChar = "";

                    //Check 1st character of value
                    if (OriginalDictionary[key] != "") startChar = OriginalDictionary[key].Substring(0, 1);

                    //if 1st char == &
                    if (startChar == "@")
                    {
                        //Key Name
                        Key_Name = OriginalDictionary[key].Split(delimiter)[1];

                        //check change in GD
                        if (Global.Dictionary[key] != OriginalDictionary[key]) Key_Value = Global.Dictionary[key];
                        else Key_Value = "";

                        //Check Records
                        SQL = "SELECT count(*) as NO_OF_RECORDS FROM [KEEP_REFER$] WHERE KEY_NAME = '" + Key_Name + "'";
                        objDS = objDB.fExecuteSelectQuery(SQL, objConn);

                        //Check if objDS is null
                        if (objDS == null)
                        {
                            //Close the connection
                            objConn.Close();
                            Global.fUpdateExecutionLog(LogType.error, "Failed to execute the query and the record set did not return any result");
                            return false;
                        }


                        //Check records
                        if (objDS.Tables[0].Rows[0]["NO_OF_RECORDS"].ToString() == "0") sSQL = "INSERT INTO [KEEP_REFER$] (KEY_NAME ,KEY_VALUE) VALUES ('" + Key_Name + "','" + Key_Value + "')";
                        else sSQL = "UPDATE [KEEP_REFER$] SET KEY_VALUE = '" + Key_Value + "' WHERE KEY_NAME = '" + Key_Name + "'";

                        objDS.Dispose();

                        //Update KR
                        //Define new ODBC Command
                        objDB.fExecuteInsertUpdateQuery(sSQL, objConn);
                        
                    }
                }
                objDS.Dispose();
                objConn.Close();
                //return success
                return true;
            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while executing fSetReferenceData function. Exception is " + e);
                objConn.Close();
                return false;
            }

        }

        ////***********************************************************************
        //// Name: 	    fUpdateTestCaseRowSkip
        //// Description:   Updates the currently executed row with X
        //// Author:        Aniket Gadre
        //// Input Params	: int row, ODC connection conn
        //// Return Values: None
        ////***********************************************************************
        //public bool fUpdateTestCaseRowSkip1(int row)
        //{
        //    DBActivities objDB = new DBActivities();
        //    OdbcConnection objConn = new OdbcConnection();

        //    try
        //    {
        //        //call the function to update Execution Log
        //        Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fUpdateTestCaseRowSkip");

        //        //Establish connection to the Calendar xls
        //        objConn = objDB.fConnectToXLS(Environment.GetEnvironmentVariable("W_MAIN_XLS"));

        //        //Handle null if connection fails
        //        if (objConn == null)
        //        {
        //            //call the function to update Execution Log
        //            Global.fUpdateExecutionLog(LogType.error, "Failed to establish connection to the calendar file. Excel path is " + Environment.GetEnvironmentVariable("W_MAIN_XLS"));

        //            return false;
        //        }


        //        //SQL
        //        string SQL = "Update [MAIN$] Set SKIP = 'X' where ID = '" + row + "'";

        //        //Call function to fire the Insert Update Query
        //        objDB.fExecuteInsertUpdateQuery(SQL, objConn);

        //        //call the function to update Execution Log
        //        Global.fUpdateExecutionLog(LogType.info, "Successfully executed fUpdateTestCaseRowSkip function");

        //        //return success
        //        //Close the conneciton
        //        objConn.Close();
        //        return true;
        //    }
        //    catch (Exception e)
        //    {
        //        //call the function to update Execution Log
        //        Global.fUpdateExecutionLog(LogType.error, "Got an exception while executing fUpdateTestCaseRowSkip function. Exception is " + e);
        //        objConn.Close();
        //        return false;
        //    }
        //}

        //***********************************************************************
        // Name: 	    fUpdateTestCaseRowSkip
        // Description:   Updates the currently executed row with X
        // Author:        Anil Agarwal
        // Input Params	: int row, string strTestName, string strResult
        // Return Values: None
        //***********************************************************************
        public bool fUpdateTestCaseRowSkip(int row, string strTestName, string strResult)
        {
            DBActivities objDB = new DBActivities();
            OdbcConnection objConn = new OdbcConnection();
            string SQL = "";
            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fUpdateTestCaseRowSkip1");

                if (strResult.ToUpper() == "FAILED")
                {
                    //Check that the test case name is not null
                    if (Global.Dictionary[strTestName] != "")
                    {
                        SQL = "Update [MAIN$] Set SKIP = 'N' where TEST_NAME ='" + Global.Dictionary[strTestName] + "' and SKIP is null";

                        //Establish connection to the Calendar xls
                        objConn = objDB.fConnectToXLS(Environment.GetEnvironmentVariable("W_MAIN_XLS"));

                        //Handle null if connection fails
                        if (objConn == null)
                        {
                            //call the function to update Execution Log
                            Global.fUpdateExecutionLog(LogType.error, "Failed to establish connection to the calendar file. Excel path is " + Environment.GetEnvironmentVariable("W_MAIN_XLS") + " while executing the query " + SQL);
                            return false;
                        }

                        //Call function to execute the query
                        if (objDB.fExecuteInsertUpdateQuery(SQL, objConn) == false)
                        {
                            Global.fUpdateExecutionLog(LogType.error, "Failed to execute the query " + SQL);
                            objConn.Close();
                            return false;
                        }
                        //Close the connection
                        objConn.Close();
                    }
                    SQL = "Update [MAIN$] Set SKIP = 'F' where ID = '" + row + "'";
                }
                else if (strResult.ToUpper() == "PASSED")
                {
                    SQL = "Update [MAIN$] Set SKIP = 'P' where ID = '" + row + "'";
                }
                else if (strResult.ToUpper() == "NO RUN")
                {
                    SQL = "Update [MAIN$] Set SKIP = 'N' where ID = '" + row + "'";
                }
                else
                {
                    SQL = "Update [MAIN$] Set SKIP = 'X' where ID = '" + row + "'";
                }

                //Establish connection to the Calendar xls
                objConn = objDB.fConnectToXLS(Environment.GetEnvironmentVariable("W_MAIN_XLS"));

                //Handle null if connection fails
                if (objConn == null)
                {
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog(LogType.error, "Failed to establish connection to the calendar file. Excel path is " + Environment.GetEnvironmentVariable("W_MAIN_XLS"));
                    return false;
                }

                //Call function to fire the Insert Update Query
                if (objDB.fExecuteInsertUpdateQuery(SQL, objConn) == false)
                {
                    Global.fUpdateExecutionLog(LogType.error, "Failed to execute the query " + SQL);
                    objConn.Close();
                    return false;
                }
                

                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.info, "Successfully executed fUpdateTestCaseRowSkip function");

                //return success
                //Close the conneciton
                objConn.Close();
                return true;
            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while executing fUpdateTestCaseRowSkip function. Exception is " + e);
                objConn.Close();
                return false;
            }
        }

        //*****************************************************************************************
        //*	Name		    : fCreateHTMLSummaryReport
        //*	Description	    : Creates the HTMLSummaryReport
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public void fCreateHTMLSummaryReport()
        {
            try
            {
                //Call function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fCreateHTMLSummaryReport wrapper");

                //Call function to create Summary Report
                Reporter.fnCreateSummaryReport();

            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Got exception while executing the fCreateHTMLSummaryReport function. Exception is " + e);
            }
        }

        //*****************************************************************************************
        //*	Name		    : fCloseHTMLSummaryReport
        //*	Description	    : Close the HTMLSummaryReport
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public void fCloseHTMLSummaryReport()
        {
            try
            {
                //Call function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fCloseHTMLSummaryReport arapper");


                //Call function to create Summary Report
                Reporter.fnCloseTestSummary();

            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Got exception while executing the fCloseHTMLSummaryReport function. Exception is " + e);
            }
        }


        //*****************************************************************************************
        //*	Name		    : fnCreateHtmlReport
        //*	Description	    : Create the HTML Report File
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public void fnCreateHtmlReport(string strFileName)
        {
            try
            {
                //Call function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fnCreateHtmlReport wrapper");

                //Call function to create Summary Report
                Reporter.fnCreateHtmlReport(strFileName);
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Got exception while executing the fnCreateHtmlReport function. Exception is " + e);
            }
        }

        //*****************************************************************************************
        //*	Name		    : fnCloseHtmlReport
        //*	Description	    : Close the HTML Report File
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public void fnCloseHtmlReport()
        {
            try
            {
                //Call function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fnCloseHtmlReport wrapper");

                //Call function to create Summary Report
                Reporter.fnCloseHtmlReport();

            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Got exception while executing the fnCloseHtmlReport function. Exception is " + e);
            }
        }

        //*****************************************************************************************
        //*	Name		    : fnWriteToHTMLOutput
        //*	Description	    : Function to Write to HTML Output
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public void fnWriteToHTMLOutput(string strDescription, string strObtainedValue, string strResult)
        {
            try
            {
                //Call function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fnWriteToHTMLOutput wrapper");

                //Call function to create Summary Report
                Reporter.fnWriteToHtmlOutput(strDescription, strObtainedValue, strResult);
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Got exception while executing the fnWriteToHTMLOutput function. Exception is " + e);
            }
        }

        //*****************************************************************************************
        //*	Name		    : fSetQCParams
        //*	Description	    : Close the HTMLSummaryReport
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public void fSetQCParams(TDAPIOLELib.TDConnection objTDConnect, string strTestSetNameTemp, string strTestSetPathTemp)
        {
            try
            {
                //Call function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fSetQCParams");

                //Set the value for the TDConection, TestSetFolderName and TestSetfolderpath
                Global.objTD = objTDConnect;
                strTestSetName = strTestSetNameTemp;
                strTestSetPath = strTestSetPathTemp;

                Global.fUpdateExecutionLog(LogType.info, "Successfully executed the fSetQCParams function");

            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Got exception while executing the fSetQCParams function. Exception is " + e);
            }
        }

        //*****************************************************************************************
        //*	Name		    : fAddTest
        //*	Description	    : Function that adds the test in QC
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public bool fAddTest(string strTestSetNameTemp, string strTestNameTemp)
        {
            //Call function to update Execution Log
            Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fAddTest");

            //Set the value for the TDConection, TestSetFolderName and TestSetfolderpath
            if (QC.fAddTest(strTestSetNameTemp, strTestNameTemp) == false) return false;

            //return
            return true;
        }

        //*****************************************************************************************
        //*	Name		    : fAttachResultsInQC
        //*	Description	    : Function that attaches result to QC for current Run
        //*	Author		    : Aniket Gadre
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public bool fAttachResultsInQC(string strTestDetails)
        {
            //Call function to update Execution Log
            Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fAttachResultsInQC");

            //Call function to attach tests
            if (QC.fAttachResultsToRun(strTestDetails) == false) return false;

            return true;
        }


        //*****************************************************************************************
        //*	Name		    : fClearSkip
        //*	Description	    : Clears the SKIP column in the data table
        //*	Author		    : Aniket Gadre
        //* Input Params	: sActionValue - The action to clear the skip field (A - Clear All, F - Clear Failed, 
        //*									No Run, and X, S - Clear Skipped and X, ABS - Clear all but Skipped) 
        //*	Return Values	: None
        //*****************************************************************************************  
        public void fClearSkip(string sActionValue)
        {
            DBActivities objDB = new DBActivities();
            OdbcConnection objConn = new OdbcConnection();

            try
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fClearSkip");

                //SQL
                string sSQL = "";

                if (sActionValue.Equals("A"))
                {
                    sSQL = "Update [MAIN$] Set SKIP = ' '";
                }

                else if (sActionValue.Equals("F"))
                {
                    sSQL = "Update [MAIN$] Set SKIP = ' ' where SKIP in ('F', 'f', 'N', 'n', 'X', 'x')";
                }

                else if (sActionValue.Equals("S"))
                {
                    sSQL = "Update [MAIN$] Set SKIP = ' ' where SKIP in ('S', 's', 'X', 'x')";
                }

                else if (sActionValue.Equals("ABS"))
                {
                    sSQL = "Update [MAIN$] Set SKIP = ' ' where SKIP not in ('S', 's')";
                }
                else
                {
                    //Call the Update Test Skip row
                    Global.fUpdateExecutionLog(LogType.error, "Update SKIP Column in Data Table - The Action: " + sActionValue + " is not valid, the valid actions to be performed are A, F, S, or ABS");
                    return;
                }
                //Establish connection to the Calendar xls
                objConn = objDB.fConnectToXLS(Environment.GetEnvironmentVariable("W_MAIN_XLS"));

                //Handle null if connection fails
                if (objConn == null)
                {
                    //call the function to update Execution Log
                    Global.fUpdateExecutionLog(LogType.error, "Failed to establish connection to the calendar file. Excel path is " + Environment.GetEnvironmentVariable("W_MAIN_XLS"));

                    return;
                }


                //Define new ODBC Command
                objDB.fExecuteInsertUpdateQuery(sSQL, objConn);

                //return success
                //Close the connection
                objConn.Close();
                return;
            }
            catch (Exception e)
            {
                //call the function to update Execution Log
                Global.fUpdateExecutionLog(LogType.error, "Got an exception while executing fClearSkip function. Exception is " + e);
                objConn.Close();
                return;
            }
        }

        //*****************************************************************************************
        //*   Name              : fDBActivities
        //*   Description       : Function that Executes the DB Activities function
        //*   Author            : Anil Agarwal
        //*   Input Params  : None
        //*   Return Values : None
        //*****************************************************************************************
        public bool fDBActivities()
        {
            try
            {
                //Call function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fDBActivities wrapper");

                //Set the value for the TDConection, TestSetFolderName and TestSetfolderpath
                if (objDB.fDBActivities() == false)
                {
                    Global.fUpdateExecutionLog(LogType.error, "Failed to Execute the fDBActivities function");
                    return false;
                }

                return true;

            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Got exception while executing the fDBActivities function. Exception is " + e);
                return false;
            }
        }

        //*****************************************************************************************
        //*   Name          : fBusPLinkExecuteUnixCommand
        //*   Description   : Function that Executes the DB Activities function
        //*   Author        : Anil Agarwal
        //*   Input Params  : None
        //*   Return Values : None
        //*****************************************************************************************
        public bool fBusPLinkExecuteUnixCommand()
        {
            try
            {
                //Call function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fBusPLinkExecuteUnixCommand wrapper");

                //Set the value for the TDConection, TestSetFolderName and TestSetfolderpath
                if (objUnix.fPLinkExecuteUnixCommand() == false)
                {
                    Global.fUpdateExecutionLog(LogType.error, "Failed to Execute the fBusPLinkExecuteUnixCommand function");
                    return false;
                }

                return true;

            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Got exception while executing the fBusPLinkExecuteUnixCommand function. Exception is " + e);
                return false;
            }
        }


        //*****************************************************************************************
        //*   Name              : fDBCheck
        //*   Description       : Function that Executes the DB Activities function
        //*   Author            : Anil Agarwal
        //*   Input Params  : None
        //*   Return Values : None
        //*****************************************************************************************
        public Scripting.Dictionary fExecuteDBCheck()
        {
            try
            {
                Dictionary<string, string> dictTemp = new Dictionary<string, string>();
                Scripting.Dictionary scriptingDict = new Scripting.Dictionary();
                //Call function to update Execution Log
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fExecuteDBCheck wrapper");

                //call the function fDBCheck
                dictTemp = objDB.fDBCheck();

                //Check if dictTemp is not null
                if (dictTemp == null)
                {
                    Global.fUpdateExecutionLog(LogType.error, "Failed to Execute the fDBCheck function");
                    return null;
                }

                //Call the function to convert the Dictionary into Scripting.Dictionary object
                if (fDictionaryToScriptingDictionary(ref scriptingDict, dictTemp) == false)
                {
                    Global.fUpdateExecutionLog(LogType.error, "Failed to Convert the Dictionary to scripting.Dictioary in fExecuteDBCheck");
                    return null;
                }
                return scriptingDict;

            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Got exception while executing the fExecuteDBCheck function. Exception is " + e);
                return null;
            }
        }

        //*****************************************************************************************
        //*	Name		    : fReallocationOfConsole
        //*	Description	    : Reallocates console at start of execution
        //*	Author		    : Anil Agarwal
        //*	Input Params	: List of Parameters from the QTP script
        //*	Return Values	: None
        //*****************************************************************************************
        public bool fReallocationOfConsole()
        {
            //if (Global.FreeConsole() == false) return false;

            if (Global.fAllocateConsole() == false) return false;


            return true;
        }


    }
}