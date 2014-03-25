using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Diagnostics;

namespace ASAP
{
    class Unix
    {

        //***********************************List of Functions in Lib ******************************
        // 1. bool fExecuteCommand(string strExe, string strCommand)
        // 2. bool fPLinkExecuteUnixCommand()
        // 3. bool fWinScpCopyFileLocal(string strUnixServer, string strUnixUser, string strUnixPassword, string strUnixPath, string strUnixFileName, string strLocalFilePath, string strLocalFileName)
        // 4. bool fWinScpCopyFileUnix(string strUnixServer, string strUnixUser, string strUnixPassword, string strUnixPath, string strUnixFileName, string strLocalFilePath, string strLocalFileName)
        //***********************************List of Functions in Lib ******************************

        //*****************************************************************************************
        //*	Name		    : fExecuteCommand
        //*	Description	    : Function to run a command form Command Prompt
        //*	Author		    : Anil Agarwal
        //*	Input Params	: string strCommand
        //*	Return Values	: bool
        //*****************************************************************************************
        public bool fExecuteCommand(string strExe, string strCommand)
        {
            Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fExecuteCommand");
            try
            {

                ProcessStartInfo procStartInfo = new ProcessStartInfo(strExe);
                procStartInfo.Arguments = strCommand;
                Process pro = new Process();
                pro.StartInfo = procStartInfo;
                pro.StartInfo.UseShellExecute = false;
                pro.StartInfo.RedirectStandardOutput = true;
                pro.StartInfo.RedirectStandardInput = true;

                pro.Start();

                pro.WaitForExit();

                Global.fUpdateExecutionLog(LogType.info, "Successfully executed the Unix Commmand " + strCommand);
                return true;
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured while executing the Unix Commmand " + strCommand);
                return false;
            }


        }

        //*****************************************************************************************
        //*	Name		    : fPLinkExecuteUnixCommand
        //*	Description	    : Function to execute a Unix command in Unix
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: bool
        //*****************************************************************************************
        public bool fPLinkExecuteUnixCommand()
        {
            Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fPLinkExecuteUnixCommand");
            //Declare variables
            string strUnixServer = "";
            string strUnixUser = "";
            string strUnixPassword = "";
            string strUnixBoxType = "";
            string strUnixPath = "";
            string strCommand = "";

            try
            {
                //Get the Unix box type
                strUnixBoxType = Global.Dictionary["UNIX_BOX_TYPE"];
                Global.fUpdateExecutionLog(LogType.debug, strUnixBoxType);

                //Set credentials depending on Unix Box Type
                strUnixServer = Environment.GetEnvironmentVariable(strUnixBoxType + "_UNIX_SERVER").Trim();
                //Global.fUpdateExecutionLog(LogType.debug, strUnixServer);
                strUnixUser = Environment.GetEnvironmentVariable(strUnixBoxType + "_UNIX_USERNAME").Trim();
                //Global.fUpdateExecutionLog(LogType.debug, strUnixUser);
                strUnixPassword = Environment.GetEnvironmentVariable(strUnixBoxType + "_UNIX_PASSWORD").Trim();
                //Global.fUpdateExecutionLog(LogType.debug, strUnixPassword);
                strUnixPath = Environment.GetEnvironmentVariable(strUnixBoxType + "_UNIX_PATH").Trim();
                //Global.fUpdateExecutionLog(LogType.debug, strUnixPath);

                //If the Unix Command contains Parameters then replace them by their values
                //Not writing this code since we all give the complete Unix command and do not provide parameters
                
                //Get the path for Plink
                string strPlinkPath = "";
                try
                {
                    strPlinkPath = Environment.GetEnvironmentVariable("PLINK_PATH");
                }
                catch (Exception e)
                {
                    Global.fUpdateExecutionLog(LogType.error, "Check if the PLINK_PATH exists in the Enviornment xls");
                    return false;
                }

                //Validate if the path contains the full path for Plink.exe
                if (!strPlinkPath.ToUpper().Contains("PLINK.EXE"))
                {
                    Global.fUpdateExecutionLog(LogType.error, "Check the PLINK_PATH given in the Enviornment xls. It should the full path including the plink.exe");
                    return false;
                }

 
                //If the PLink Path contains " " then add double quotes to it
                if (strPlinkPath.Contains(" "))
                {
                    strPlinkPath = "\"" + strPlinkPath + "\"";
                }

                //Get the Unix file name, Local file name and Local file path from the Dictionary calendar
                string strUnixFileName = Global.Dictionary["UNIX_FILE_NAME"];
                string strLocalFilePath = Global.Dictionary["LOCAL_FILE_PATH"];
                string strLocalFileName = Global.Dictionary["LOCAL_FILE_NAME"];

                //Check the Unix file path for a "\" in the end
                if (!strUnixPath.EndsWith("/"))
                {
                    strUnixPath = strUnixPath + "/";
                }

                //Create the command
                strCommand = Global.Dictionary["COMMAND"];
                strCommand = strUnixUser + "@" + strUnixServer + " -pw " + strUnixPassword + " . ./.profile;" + strCommand + " >" + strUnixPath + strUnixFileName;
                Global.fUpdateExecutionLog(LogType.debug, "Executing command: " + strCommand);
                //Call the function to call the command
                if (fExecuteCommand(strPlinkPath, strCommand) == false)
                {
                    return false;
                }

                //call the function to FTP the Output file from Unix box to local system
                if (fWinScpCopyFileLocal(strUnixServer, strUnixUser, strUnixPassword, strUnixPath, strUnixFileName, strLocalFilePath, strLocalFileName) == false)
                {
                    return false;
                }


                Global.fUpdateExecutionLog(LogType.info, "Successfully executed the fPLinkExecuteUnixCommand Function");
                return true;
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured while executing the Unix Commmand " + strCommand);
                return false;
            }
        }

        //*****************************************************************************************
        //*	Name		    : fWinScpCopyFileLocal
        //*	Description	    : Function to Copy a file from Unix into local using WinScp
        //*	Author		    : Anil Agarwal
        //*	Input Params	: string strUnixServer, string strUnixUser, string strUnixPassword
        //*                   string strUnixPath, string strUnixFileName, string strLocalFilePath, string strLocalFileName                 
        //*	Return Values	: bool
        //*****************************************************************************************
        public bool fWinScpCopyFileLocal(string strUnixServer, string strUnixUser, string strUnixPassword, string strUnixPath, string strUnixFileName, string strLocalFilePath, string strLocalFileName)
        {
            Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fWinScpCopyFileLocal");
            //Declare variables
            string strCommand = "";
            try
            {
                //Get the path for Plink
                string sWinSCPPath = "";
                try
                {
                    sWinSCPPath = Environment.GetEnvironmentVariable("WINSCP_PATH");
                }
                catch (Exception e)
                {
                    Global.fUpdateExecutionLog(LogType.error, "Check if the WINSCP_PATH exists in the Enviornment xls");
                    return false;
                }

                //Validate if the path contains the full path for Plink.exe
                if (!sWinSCPPath.ToUpper().Contains("WINSCP.COM"))
                {
                    Global.fUpdateExecutionLog(LogType.error, "Check the WINSCP_PATH given in the Enviornment xls. It should the full path including the winscp.com");
                    return false;
                }

                //Get the path for Plink
                sWinSCPPath = Environment.GetEnvironmentVariable("WINSCP_PATH").Trim();

                //If the PLink Path contains " " then add double quotes to it
                if (sWinSCPPath.Contains(" "))
                {
                    sWinSCPPath = "\"" + sWinSCPPath + "\"";
                }

                //Check the Local file path for a "\" in the end
                if (!strLocalFilePath.EndsWith("\\"))
                {
                    strLocalFilePath = strLocalFilePath + "\\";
                }

                //Check the Unix file path for a "\" in the end
                if (!strUnixPath.EndsWith("/"))
                {
                    strUnixPath = strUnixPath + "/";
                }

                //Create the command
                strCommand = "/Command" + " \"open sftp://" + strUnixUser + ":" + strUnixPassword + "@" + strUnixServer + "\" " + "\"get " + strUnixPath + strUnixFileName + " " + strLocalFilePath + strLocalFileName + "\" " + "\"exit\"";

                //Call the function to call the command
                if (fExecuteCommand(sWinSCPPath, strCommand) == false)
                {
                    return false;
                }

                Global.fUpdateExecutionLog(LogType.info, "Successfully FTP the Unix file " + strUnixFileName + " in the Local path at " + strLocalFilePath + " with file name " + strLocalFileName);
                return true;
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured while trying to ftp the file " + strUnixFileName);
                return false;
            }

        }


        //*****************************************************************************************
        //*	Name		    : fWinScpCopyFileUnix
        //*	Description	    : Function to Copy a file from local into Unix using WinScp
        //*	Author		    : Anil Agarwal
        //*	Input Params	: string strUnixServer, string strUnixUser, string strUnixPassword
        //*                   string strUnixPath, string strUnixFileName, string strLocalFilePath, string strLocalFileName                 
        //*	Return Values	: bool
        //*****************************************************************************************
        public bool fWinScpCopyFileUnix(string strUnixServer, string strUnixUser, string strUnixPassword, string strUnixPath, string strUnixFileName, string strLocalFilePath, string strLocalFileName)
        {
            Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fWinScpCopyFileUnix");
            //Declare variables
            string strCommand = "";
            try
            {
                //Get the path for Plink
                string sWinSCPPath = Environment.GetEnvironmentVariable("WINSCP_PATH").Trim();

                //If the PLink Path contains " " then add double quotes to it
                if (sWinSCPPath.Contains(" "))
                {
                    sWinSCPPath = "\"" + sWinSCPPath + "\"";
                }


                //If the PLink Path contains " " then add double quotes to it
                if (sWinSCPPath.Contains(" "))
                {
                    sWinSCPPath = "\"" + sWinSCPPath + "\"";
                }

                //Check the Local file path for a "\" in the end
                if (!strLocalFilePath.EndsWith("\\"))
                {
                    strLocalFilePath = strLocalFilePath + "\\";
                }

                //Check the Unix file path for a "\" in the end
                if (!strUnixPath.EndsWith("/"))
                {
                    strUnixPath = strUnixPath + "/";
                }


                //Create the command
                strCommand = "/Command" + " \"open sftp://" + strUnixUser + ":" + strUnixPassword + "@" + strUnixServer + "\" " + "\"put " + strLocalFilePath + strLocalFileName + " " + strUnixPath + strUnixFileName + "\" " + "\"exit\"";

                //Call the function to call the command
                if (fExecuteCommand(sWinSCPPath, strCommand) == false)
                {
                    return false;
                }

                Global.fUpdateExecutionLog(LogType.info, "Successfully FTP the Local file " + strLocalFileName + " in Unix at path  " + strUnixPath + " and  file name " + strUnixFileName);
                return true;
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured while trying to ftp the file " + strLocalFileName);
                return false;
            }

        }


    }
}
