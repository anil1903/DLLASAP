using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ASAP
{

    public enum LogType { info, error, debug };


    public static class Global
    {
        private const UInt32 StdOutputHandle = 0xFFFFFFF5;

        [DllImport("kernel32.dll")]
        public static extern bool AllocConsole();

        [DllImport("kernel32.dll")]
        public static extern bool FreeConsole();

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetStdHandle(UInt32 nStdHandle);

        [DllImport("kernel32.dll")]
        private static extern void SetStdHandle(UInt32 nStdHandle, IntPtr handle);


        public static Dictionary<string, string> Dictionary = new Dictionary<string, string>();
        public static TDAPIOLELib.TDConnection objTD = null;

        //*****************************************************************************************
        //*	Name		    : fUpdateExecutionLog
        //*	Description	    : Update the Execution Log file to keep a track of all execution and failures
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public static void fUpdateExecutionLog(string strLog)
        {
            try
            {
                //Open the execution Log File
                StreamWriter streamWriter = new StreamWriter(Environment.GetEnvironmentVariable("EXECUTION_LOG_FILE_PATH"), true);

                //Writing the log in the execution log file
                using (streamWriter)
                {
                    streamWriter.WriteLine(">>>> " + strLog);
                }

                //AttachConsole(ATTACH_PARENT_PROCESS);
                //AllocConsole();
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine(">>>> " + strLog);

            }
            catch (Exception e)
            {
                Console.WriteLine("Exception " + e + " occured while updating log");
            }
        }

        //*****************************************************************************************
        //*	Name		    : fUpdateExecutionLog
        //*	Description	    : Update the Execution Log file to keep a track of all execution and failures
        //*	Author		    : Anil Agarwal
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public static void fUpdateExecutionLog(LogType lgType, string strLog)
        {
            try
            {
                //Open the execution Log File
                StreamWriter streamWriter = new StreamWriter(Environment.GetEnvironmentVariable("EXECUTION_LOG_FILE_PATH"), true);
                string strLogType;

                //prefix log type
                if (Convert.ToInt32(lgType) == 0)
                {
                    strLogType = "Info: ";
                    Console.ForegroundColor = ConsoleColor.Cyan;
                }
                else if (Convert.ToInt32(lgType) == 1)
                {
                    strLogType = "Error: ";
                    Console.ForegroundColor = ConsoleColor.Red;
                }
                else
                {
                    strLogType = "Debug: ";
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                }

                //Writing the log in the execution log file
                using (streamWriter)
                {
                    streamWriter.WriteLine(strLogType + strLog);
                }

                //AttachConsole(ATTACH_PARENT_PROCESS);
                //AllocConsole();

                Console.WriteLine(strLogType + strLog);

                //Freeing console after writing 
                //FreeConsole();

            }
            catch (Exception e)
            {
                Console.WriteLine("Exception " + e + "occured while updating the log");
            }
        }


        //*****************************************************************************************
        //*	Name		    : fAllocateConsole
        //*	Description	    : Allocates a console to QTP process
        //*	Author		    : Aniket Gadre
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public static bool fAllocateConsole()
        {
            try
            {

                AllocConsole();

                // stdout's handle seems to always be equal to 7
                IntPtr defaultStdout = new IntPtr(7);
                IntPtr currentStdout = GetStdHandle(StdOutputHandle);

                if (currentStdout != defaultStdout)
                    // reset stdout
                    SetStdHandle(StdOutputHandle, defaultStdout);

                // reopen stdout
                TextWriter writer = new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true };
                Console.SetOut(writer);
                fUpdateExecutionLog(LogType.info, "Allocated Console");

            }
            catch (Exception e)
            {
                fUpdateExecutionLog(LogType.error, "Failed allocating Console. Exception occured : " + e);
                return false;
            }

            return true;
        }


        //*****************************************************************************************
        //*	Name		    : fAllocateConsole
        //*	Description	    : Allocates a console to QTP process
        //*	Author		    : Aniket Gadre
        //*	Input Params	: None
        //*	Return Values	: None
        //*****************************************************************************************
        public static bool fFreeConsole()
        {
            try
            {

                FreeConsole();
            }
            catch (Exception e)
            {
                fUpdateExecutionLog(LogType.error, "Failed freeing Console. Exception occured : " + e);
                //return false;
            }

            return true;
        }
    }
}
