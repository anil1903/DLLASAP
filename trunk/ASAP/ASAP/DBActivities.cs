using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;

namespace ASAP
{
    public class DBActivities
    {

        //******* Methods in this class **********//
        //** OdbcConnection fConnectToXLS(string XLSPath) - ##DONE##
        //** OdbcConnection fConnectToDB(string strServer, string strUserName, string strPassword) ##DONE##
        //** DataSet fExecuteSelectQuery(string SQL,OdbcConnection conn) ##DONE##
        //** void fExecuteInsertUpdateQuery(string SQL,OdbcConnection conn) ##DONE##
        //** string fGetReferenceVerificationData(string strParamName) ##DONE##
        //** bool fSetReferenceVerificationData(string strParamNames, Dictionary<string, string> outDict) ##DONE##
        //** string fReplaceSpecialParameterInSQL(string strSQL) ##DONE##
        //** bool fDBActivities()
        //** Dictionary<string, string> fDBCheck()
        //** int fGetColumnName(DataTable DT, string strColumnName)
        //** bool fDBCompareArray(string[] arrExpRes, Dictionary<string,string> octDictTemp) ##DONE##
        //** bool fValidateComparisonValues(string str1, string str2) ##DONE##
        //******* Methods in this class **********//


        //*****************************************************************************************
        //*	Name		    : fConnectToXLS
        //*	Description	    : Function to make an ODBC connection to the given xls
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string XLSPath - Path of the xls to be connected to
        //*	Return Values	: OdbcConnection obj
        //*****************************************************************************************
        public OdbcConnection fConnectToXLS(string XLSPath)
        {
            //log
            Global.fUpdateExecutionLog(LogType.debug, "Executing function fConnectToXLS");

            //Create a new ODBC instance
            OdbcConnection conn = new OdbcConnection();

            try
            {
                //Conn string 1
                conn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DriverId=790;FIL=Excel 12.0;Dbq= " + XLSPath + ";ReadOnly=0;";

                //log
                Global.fUpdateExecutionLog(LogType.debug, "Using Connection string: " + conn.ConnectionString);

                //Open connection
                conn.Open();
            }

            //Handle connection exceptions
            catch (Exception e)
            {

                try
                {

                    //Conn string 1
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + XLSPath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;READONLY=FALSE\";";

                    //log
                    Global.fUpdateExecutionLog(LogType.debug, "Using Connection string: " + conn.ConnectionString);

                    //Open connection
                    conn.Open();
                }

                catch (Exception e1)
                {
                    try
                    {

                        //Conn string 1
                        conn.ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" + XLSPath + "; Mode=ReadWrite; Extended Properties=\"Excel 8.0; HDR=Yes;READONLY=FALSE\"";

                        //log
                        Global.fUpdateExecutionLog(LogType.debug, "Using Connection string: " + conn.ConnectionString);

                        //Open connection
                        conn.Open();
                    }

                    catch (Exception e2)
                    {
                        return null;
                    }
                }
            }


            //log
            Global.fUpdateExecutionLog(LogType.info, "DB Connection to xls " + XLSPath + " successfull");
            return conn;
        }

        //*****************************************************************************************
        //*	Name		    : fConnectToDB
        //*	Description	    : Function to make connection to DB
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string strDBType - SQL Query string                 
        //*	Return Values	: OdbcConnection object
        //*****************************************************************************************
        public OleDbConnection fConnectToDB(string strDBType)
        {

            //log
            Global.fUpdateExecutionLog(LogType.debug, "Executing function fConnectToDB");

            //Declare variables
            string strServer = "";
            string strUserName = "";
            string strPassword = "";

            try
            {

                //Set credentials depending on DB Type
                strServer = Environment.GetEnvironmentVariable(strDBType + "_DB_SERVER").Trim();
                strUserName = Environment.GetEnvironmentVariable(strDBType + "_DB_USERNAME").Trim();
                strPassword = Environment.GetEnvironmentVariable(strDBType + "_DB_PASSWORD").Trim();


                //Create New Connection
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "Provider=oraOLEDB.Oracle ;Data Source =" + strServer + ";User Id=" + strUserName + ";Password=" + strPassword;

                //Open Connection
                conn.Open();

                //Connection Success
                Global.fUpdateExecutionLog(LogType.info, "Successfully connected to " + strDBType + " DB with Credentials " + strServer + " " + strUserName + " " + strPassword);
                return conn;
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured while connecting to DB of type " + strDBType + " with Credentials " + strServer + " " + strUserName + " " + strPassword);
                return null;
            }

        }

        //*****************************************************************************************
        //*	Name		    : fExecuteSelectQuery
        //*	Description	    : Function to execute an query
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string SQL - SQL Query string
        //*                   OdbcConnection conn - ODBC connection object
        //*	Return Values	: Dataset object having the results
        //*****************************************************************************************
        public DataSet fExecuteSelectQuery(string SQL, OdbcConnection conn)
        {
            //log
            Global.fUpdateExecutionLog(LogType.debug, "Executing function fExecuteSelectQuery");

            //display query
            //Global.fUpdateExecutionLog(LogType.info, "Executing query " + SQL);

            try
            {
                //Open a recordet by executing the query
                OdbcDataAdapter adapter = new OdbcDataAdapter(SQL, conn);
                DataSet ds = new DataSet();
                adapter.Fill(ds);

                //Log
                Global.fUpdateExecutionLog(LogType.info, "Successfully executed query " + SQL);
                return ds;
            }
            //catch any exception thrown
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured while executing query " + SQL);
                return null;
            }
        }

        //*****************************************************************************************
        //*	Name		    : fExecuteSelectQuery
        //*	Description	    : Function to execute an query
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string SQL - SQL Query string
        //*                   OdbcConnection conn - ODBC connection object
        //*	Return Values	: Dataset object having the results
        //*****************************************************************************************
        public DataSet fExecuteSelectQuery(string SQL, OleDbConnection conn)
        {
            //log
            Global.fUpdateExecutionLog(LogType.debug, "Executing function fExecuteSelectQuery");

            try
            {
                //Open a recordet by executing the query
                OleDbDataAdapter adapter = new OleDbDataAdapter(SQL, conn);
                DataSet ds = new DataSet();
                adapter.Fill(ds);

                //Log
                Global.fUpdateExecutionLog(LogType.info, "Successfully executed query " + SQL);
                return ds;
            }
            //catch any exception thrown
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured while executing query " + SQL);
                return null;
            }
        }


        //*****************************************************************************************
        //*	Name		    : fExecuteInsertUpdateQuery
        //*	Description	    : Function to execute an query
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string SQL - SQL Query string
        //*                   OdbcConnection conn - ODBC connection object
        //*	Return Values	: True - Success, False - Failure
        //*****************************************************************************************
        public bool fExecuteInsertUpdateQuery(string SQL, OdbcConnection conn)
        {
            //log
            Global.fUpdateExecutionLog(LogType.debug, "Executing function fExecuteInsertUpdateQuery");

            try
            {
                //Define new ODBC Command
                OdbcCommand cmd = new OdbcCommand(SQL, conn);
                int iRowsAffected = cmd.ExecuteNonQuery();

                //Check number of rows
                if (iRowsAffected == 0)
                {
                    Global.fUpdateExecutionLog(LogType.info, "No records were inserted or updated by query " + SQL);
                    return false;
                }

                //log
                Global.fUpdateExecutionLog(LogType.info, "Successfully executed query " + SQL);
                return true;
            }

            //catch any exception thrown
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured while executing query " + SQL);
                return false;
            }
        }

        //*****************************************************************************************
        //*	Name		    : fExecuteInsertUpdateQuery
        //*	Description	    : Function to execute an query
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string SQL - SQL Query string
        //*                   OdbcConnection conn - ODBC connection object
        //*	Return Values	: True - Success, False - Failure
        //*****************************************************************************************
        public bool fExecuteInsertUpdateQuery(string SQL, OleDbConnection conn)
        {
            //log
            Global.fUpdateExecutionLog(LogType.debug, "Executing function fExecuteInsertUpdateQuery");

            try
            {
                //Define new ODBC Command
                OleDbCommand cmd = new OleDbCommand(SQL, conn);
                int iRowsAffected = cmd.ExecuteNonQuery();

                //Check number of rows
                if (iRowsAffected == 0)
                {
                    Global.fUpdateExecutionLog(LogType.error, "No records were inserted or updated by query " + SQL);
                    return false;
                }

                //log
                Global.fUpdateExecutionLog(LogType.info, "Successfully executed query " + SQL);
                return true;
            }

            //catch any exception thrown
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured while executing query " + SQL);
                return false;
            }
        }


        //*****************************************************************************************
        //*	Name		    : fGetReferenceVerificationData
        //*	Description	    : Function to get the data from KEEP_REFER sheet corresponding to passed param
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string strParamName - Name of Key       
        //*	Return Values	: Value against Key - Success, Null String - Failure
        //*****************************************************************************************
        public string fGetReferenceVerificationData(string strParamName)
        {

            //log
            Global.fUpdateExecutionLog(LogType.debug, "Executing function fGetReferenceVerificationData");

            //Get the Calendar Excel Sheet
            string strCalPath = Environment.GetEnvironmentVariable("W_MAIN_XLS");

            //SQL
            string strSQL = "Select KEY_VALUE From [KEEP_REFER$] Where KEY_NAME = '" + strParamName + "'";

            try
            {

                //ODB COnnection
                OdbcConnection conn = fConnectToXLS(strCalPath);
                if (conn == null)
                {
                    Global.fUpdateExecutionLog(LogType.error, "Failed to establish connection to the xls  " + strCalPath);
                    return "";
                }


                //Execute Query
                DataSet DS = fExecuteSelectQuery(strSQL, conn);
                if (DS == null)
                {
                    Global.fUpdateExecutionLog(LogType.error, "The query did not return any result. SQL is " + strSQL);
                    conn.Close();
                    return "";
                }

                //Check if records are populated
                if (DS.Tables[0].Rows.Count == 0)
                {
                    Global.fUpdateExecutionLog(LogType.error, "No records fetched by query " + strSQL);
                    conn.Close();
                    return "";
                }
                else
                {
                    string retVal = DS.Tables[0].Rows[0].ItemArray[0].ToString();
                    if (retVal == null || retVal == "")
                    {
                        Global.fUpdateExecutionLog(LogType.info, "Value for param " + strParamName + " is empty in KEEP_REFER sheet");
                    }

                    conn.Close();
                    return retVal;
                }
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured while executing query " + strSQL);
                return "";
            }

        }

        //*****************************************************************************************
        //*	Name		    : fSetReferenceVerificationData
        //*	Description	    : Function to set the data in KEEP_REFER sheet corresponding to passed params
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string strParamNames - Names of Key       
        //*	Return Values	: Value against Key - Success, Null String - Failure
        //*****************************************************************************************
        public bool fSetReferenceVerificationData(string strParamNames, Dictionary<string, string> outDict)
        {

            //log
            Global.fUpdateExecutionLog(LogType.debug, "Executing function fSetReferenceVerificationData");

            //Define Temp Vars
            string strTempParamName, strTempParamValue;

            //Get the Calendar Excel Sheet
            string strCalPath = Environment.GetEnvironmentVariable("W_MAIN_XLS");

            //Arrays
            string[] arrParamName;
            string[] arrParamValue;

            //Delimiters
            char[] delimiters = { ';', ',' };
            string[] frmDelimiters = { "FROM" };
            string[] selDelimiters = { "SELECT" };
            //ODB COnnection
            OdbcConnection conn = new OdbcConnection();

            try
            {

                //Check if strParamNames is not null
                if (strParamNames.Trim() != "")
                {
                    arrParamName = strParamNames.Split(delimiters);
                }
                else
                {
                    // arrParamName = (Dictionary["DBSQL"].ToUpper().Split(frmDelimiters))[0].Trim().Split(selDelimiters)[1].Trim().Split(delimiters);
                    arrParamName = ((Global.Dictionary["DBSQL"].ToUpper().Split(frmDelimiters, StringSplitOptions.RemoveEmptyEntries))[0].Trim().Split(selDelimiters, StringSplitOptions.RemoveEmptyEntries))[1].Trim().Split(delimiters);
                }

                //fetch arr of values
                arrParamValue = outDict.Values.ToArray();


                conn = fConnectToXLS(strCalPath);

                if (conn == null)
                {
                    Global.fUpdateExecutionLog(LogType.error, "Failed to establish connection to the xls  " + strCalPath);
                    return false;
                }


                //loop through al paramValues
                int cnt = arrParamValue.Count();

                for (int z = 0; z < cnt; z++)
                {
                    strTempParamName = arrParamName[z];
                    strTempParamValue = arrParamValue[z];

                    //Query
                    string strSQL = "Select count(*) as NO_OF_RECORDS from [KEEP_REFER$]  where KEY_NAME = '" + strTempParamName + "'";

                    //Fetch records
                    DataSet DS = fExecuteSelectQuery(strSQL, conn);
                    if (DS == null)
                    {
                        Global.fUpdateExecutionLog(LogType.error, "The query did not return any result. SQL is " + strSQL);
                        conn.Close();
                        return false;
                    }

                    //Row count
                    int iRowCnt = Convert.ToInt32(DS.Tables[0].Rows[0].ItemArray[0].ToString());

                    //If rowcount is 0 insert , else update
                    if (iRowCnt == 0)
                    {
                        strSQL = "Insert into [KEEP_REFER$] (KEY_NAME, KEY_VALUE) values ('" + strTempParamName + "','" + strTempParamValue + "')";
                    }
                    else
                    {
                        strSQL = "Update [KEEP_REFER$] Set KEY_VALUE = '" + strTempParamValue + "' Where KEY_NAME = '" + strTempParamName + "'";
                    }

                    //Execute QUery
                    if (fExecuteInsertUpdateQuery(strSQL, conn) == false)
                    {
                        Global.fUpdateExecutionLog(LogType.error, "Executing query " + strSQL + " failed");
                        conn.Close();
                        return false;
                    }

                    //Put in Dictionary
                    if (Global.Dictionary.ContainsKey(strTempParamName)) Global.Dictionary[strTempParamName] = strTempParamValue;
                    else Global.Dictionary.Add(strTempParamName, strTempParamValue);

                }

                conn.Close();
                return true;
            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured in function fSetReferenceVerificationData");
                conn.Close();
                return false;
            }

        }

        //*****************************************************************************************
        //*	Name		    : fReplaceSpecialParameterInSQL
        //*	Description	    : Function to replace param with its value in SQL from Keep_Refer shee
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string strSQL- Query from which params are to be replaced   
        //*	Return Values	: string strSQL - Updated Query
        //*****************************************************************************************
        public string fReplaceSpecialParameterInSQL(string strSQL)
        {
            //log
            Global.fUpdateExecutionLog(LogType.debug, "Executing function fReplaceSpecialParameterInSQL");

            //Declare variables
            string strParamName, strParamValue;

            //Delimiter
            string[] startDelimiter = { "<&" };
            string[] endDelimiter = { ">" };


            try
            {

                //Check if any param needs to be replaced
                while (strSQL.Contains("<&"))
                {
                    //Get the param Name
                    strParamName = (strSQL.Split(startDelimiter, StringSplitOptions.RemoveEmptyEntries))[1].Split(endDelimiter, StringSplitOptions.RemoveEmptyEntries)[0];

                    //Fetch param value corresponding to param name
                    strParamValue = fGetReferenceVerificationData(strParamName);

                    //check if value is empty
                    if (strParamValue == "")
                    {
                        Global.fUpdateExecutionLog(LogType.error, "No value specified for parameter " + strParamName + " in KEEP_REFER sheet. This param is used in SQL " + strSQL);
                        return "";
                    }

                    //replace value
                    strSQL = strSQL.Replace("<&" + strParamName + ">", strParamValue);
                }


                return strSQL;

            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Exception " + e + " occured in function fReplaceSpecialParameterInSQL() while replacing special params in SQL " + strSQL);
                return "";
            }

        }

        //*****************************************************************************************
        //*	Name		    : fDBActivities
        //*	Description	    : FUnction to execute all DB Activities
        //*	Author		    : Aniket Gadre
        //*	Input Params	: None
        //*	Return Values	: Boolean True on Success, False on Failure
        //*****************************************************************************************
        public bool fDBActivities()
        {
            //log
            Global.fUpdateExecutionLog(LogType.debug, "Executing function fDBActivities");

            string strCommonFile, strSQL;
            strSQL = "";

            //Setting flag
            bool fDBActivitiesFlag = false;

            Dictionary<string, string> outGBL = new Dictionary<string, string>();
            Dictionary<string, string> outResult = new Dictionary<string, string>();

            //Set DB SQL SHeet
            if (Environment.GetEnvironmentVariable("DB_SQL_FROM_DATATABLE") == "" || Environment.GetEnvironmentVariable("DB_SQL_FROM_DATATABLE") == null)
            {
                Environment.SetEnvironmentVariable("DB_SQL_FROM_DATATABLE", "N");

                //log
                Global.fUpdateExecutionLog(LogType.debug, "DB_SQL_FROM_DATATABLE is set to N");
            }

            if (Environment.GetEnvironmentVariable("DB_SQL_FROM_DATATABLE").ToUpper() == "N")
            {
                strCommonFile = Environment.GetEnvironmentVariable("W_COMMON_XLS");
            }
            else
            {
                strCommonFile = Environment.GetEnvironmentVariable("W_MAIN_XLS");
            }

            //log
            Global.fUpdateExecutionLog(LogType.info, "DB SQL Excel path has been set to " + strCommonFile);

            //log
            Global.fUpdateExecutionLog(LogType.info, "Auto DP is " + Environment.GetEnvironmentVariable("AUTO_DP"));

            //Check for AutoDP
            if (Environment.GetEnvironmentVariable("AUTO_DP") != "")
            {
                if (Environment.GetEnvironmentVariable("AUTO_DP").Trim().ToUpper() == "Y")
                {
                    //strSQL1 = "Select count(*) as ROW_COUNT from [DB_SQL$] where GROUP_NAME = 'DP' and QUERY_TYPE = 'DP'";
                    strSQL = "Select * from [DB_SQL$] where GROUP_NAME = 'DP' and QUERY_TYPE = 'DP'";

                    //Add the AUTO_DP value to the dictionary object
                    if (Global.Dictionary.ContainsKey("AUTO_DP"))
                    {
                        Global.Dictionary["AUTO_DP"] = "TRUE";
                    }
                    else
                    {
                        Global.Dictionary.Add("AUTO_DP", "TRUE");
                    }


                    Environment.SetEnvironmentVariable("AUTO_DP", "N");
                }
            }

            //log
            Global.fUpdateExecutionLog(LogType.info, "Auto DP is: " + Environment.GetEnvironmentVariable("AUTO_DP"));
            if (Global.Dictionary.ContainsKey("GROUP_NAME")) Global.fUpdateExecutionLog(LogType.info, "GROUP_NAME is: " + Global.Dictionary["GROUP_NAME"].Trim());

            //Check for Group Name
            if (Environment.GetEnvironmentVariable("AUTO_DP").Trim().ToUpper() != "Y" && Global.Dictionary.ContainsKey("GROUP_NAME") && Global.Dictionary["GROUP_NAME"].Trim() != "")
            {

                //log
                Global.fUpdateExecutionLog(LogType.info, "Group Name: " + Global.Dictionary["GROUP_NAME"]);

                if (Global.Dictionary["GROUP_NAME"].Trim() == "")
                {
                    Global.fUpdateExecutionLog(LogType.info, "No value provided for GROUP_NAME parameter in MAIN SHEET");
                    return false;
                }
                else
                {
                    //strSQL1 = "Select count(*) as ROW_COUNT from [DB_SQL$] where GROUP_NAME = '" + Dictionary["GROUP_NAME"].Trim() + "'";
                    strSQL = "Select * from [DB_SQL$] where GROUP_NAME = '" + Global.Dictionary["GROUP_NAME"].Trim() + "'";
                }
            }



            //Check for PK
            else if (Environment.GetEnvironmentVariable("AUTO_DP").Trim().ToUpper() != "Y" && Global.Dictionary.ContainsKey("PK") && Global.Dictionary["PK"].Trim() != "")
            {
                //log
                Global.fUpdateExecutionLog(LogType.info, "PK: " + Global.Dictionary.ContainsKey("PK"));

                if (Global.Dictionary["PK"].Trim() == "")
                {
                    Global.fUpdateExecutionLog(LogType.error, "No value provided for PK parameter in MAIN SHEET");
                    return false;
                }
                else
                {
                    //strSQL1 = "Select count(*) as ROW_COUNT from [DB_SQL$] where PK = '" + Dictionary["PK"].Trim() + "'";
                    strSQL = "Select * from [DB_SQL$] where PK = '" + Global.Dictionary["PK"].Trim() + "'";
                }
            }else
			{
				Global.fUpdateExecutionLog(LogType.error, "No value provided for PK or GROUP_NAME parameters");
                return false;
			}
            //Connect to XLS
            //ODB COnnection
            OdbcConnection conn = fConnectToXLS(strCommonFile);
            if (conn == null)
            {
                Global.fUpdateExecutionLog(LogType.error, "Failed to establish connection to the xls " + strCommonFile);
                return false;
            }

            //Execute Query
            DataSet DS = fExecuteSelectQuery(strSQL, conn);
            if (DS == null)
            {
                //Close connection and return false;
                Global.fUpdateExecutionLog(LogType.error, "Query did not return any result. Query is " + strSQL);
                conn.Close();
                return false;
            }

            //Check no of rows
            if (DS.Tables[0].Rows.Count == 0)
            {
                Global.fUpdateExecutionLog(LogType.error, " No records fetched for query " + strSQL);
                conn.Close();
                return false;
            }
            else
            {
                //Get row count
                DataTable DT = DS.Tables[0];
                int iRecCnt = DT.Rows.Count;
                int iColCnt = DT.Columns.Count;

                string strFieldName, strFieldValue;

                int iPKIndex = 0;
                //int iIDIndex = 0;



                ///Loop through rows
                for (int m = 0; m < iRecCnt; m++)
                {
                    //Loop through fields
                    for (int n = 0; n < iColCnt; n++)
                    {
                        //Fetch Filed Name and Value
                        strFieldName = DT.Columns[n].ColumnName;
                        strFieldValue = DT.Rows[m].ItemArray[n].ToString();

                        //Check is field name is PK
                        if (strFieldName.ToUpper().Trim() == "PK")
                        {
                            //put value in outGBL
                            outGBL.Add(strFieldName + "_" + iPKIndex, strFieldValue);
                            iPKIndex++;

                            if (!Global.Dictionary.ContainsKey("PK")) Global.Dictionary.Add("PK", strFieldValue);
                            else Global.Dictionary["PK"] = strFieldValue;
                        }


                        //######## This code seems to be an overhead and may not be required ############
                        /*//Check is field name is ID
                        if (strFieldName.ToUpper().Trim() == "ID")
                        {
                            //put value in outGBL
                            outGBL.Add(strFieldName + "_" + iIDIndex, strFieldValue);
                            iIDIndex++;
                            Dictionary.Add("XL_COMMON_ID", strFieldValue);
                        }*/
                    }

                    //Call function fDBCheck
                    outResult = fDBCheck();

                    //Validate outResult
                    if (outResult == null)
                    {
                        Global.fUpdateExecutionLog(LogType.error, "fDBCheck function returned Null");
                        fDBActivitiesFlag = false;
                    }
                    else fDBActivitiesFlag = true;

                }

                //Closing Connection
                conn.Close();

            }

            //Dictionary["XL_COMMON_ID"] = "";  ###### OVERHEAD CODE ########

            //Check Flag Value
            if (!fDBActivitiesFlag)
            {
                Global.fUpdateExecutionLog(LogType.error, "fDBActivities function didn't execute successfully");
                return false;
            }

            Global.fUpdateExecutionLog(LogType.info, "fDBActivities function executed successfully");
            return true;
        }

        //*****************************************************************************************
        //*	Name		    : fDBCheck
        //*	Description	    : FUnction to validate SQL results
        //*	Author		    : Aniket Gadre
        //*	Input Params	: None
        //*	Return Values	: Boolean True on Success, False on Failure
        //*****************************************************************************************
        public Dictionary<string, string> fDBCheck()
        {
            OdbcConnection conn = null;
            //Create Connection object for DB
            OleDbConnection objConn = new OleDbConnection();
            try
            {
                string strCommonFile, strDBType, strDBSyncTime, strSaveToKR, strSaveParamName, strExpectedResult, strQuery;
                bool resDBCheck = false;
                bool ExpectedResFlag = false;
                bool Flag = false;
                Dictionary<string, string> outDictionary = new Dictionary<string, string>();
                Global.fUpdateExecutionLog(LogType.debug, "Entered Function: fDBCheck");

                //Set DB SQL SHeet
                //Set DB SQL SHeet
                if (Environment.GetEnvironmentVariable("DB_SQL_FROM_DATATABLE") == "" || Environment.GetEnvironmentVariable("DB_SQL_FROM_DATATABLE") == null)
                {
                    Environment.SetEnvironmentVariable("DB_SQL_FROM_DATATABLE", "N");
                }
                //   Global.fUpdateExecutionLog("Value is " + Environment.GetEnvironmentVariable("DB_SQL_FROM_DATATABLE"));
                if (Environment.GetEnvironmentVariable("DB_SQL_FROM_DATATABLE").ToUpper() == "N")
                {
                    strCommonFile = Environment.GetEnvironmentVariable("W_COMMON_XLS");
                }
                else
                {
                    strCommonFile = Environment.GetEnvironmentVariable("W_MAIN_XLS");
                }

                //form query
                string strSQL = "Select * from [DB_SQL$] where PK = '" + Global.Dictionary["PK"] + "'";

                //Connect to XLS
                conn = fConnectToXLS(strCommonFile);
                if (conn == null)
                {
                    Global.fUpdateExecutionLog(LogType.error, "Failed to establish connection to the xls " + strCommonFile);
                    return null;
                }


                //Execute Query
                DataSet DS = fExecuteSelectQuery(strSQL, conn);
                if (DS == null)
                {
                    //Close connection and return false;
                    Global.fUpdateExecutionLog(LogType.error, "Query did not return any result. SQL is " + strSQL);
                    conn.Close();
                    return null;
                }

                //Get row count
                DataTable DT = DS.Tables[0];
                int iRowCount = DT.Rows.Count;

                //Check record count
                if (iRowCount == 0)
                {
                    Global.fUpdateExecutionLog(LogType.error, "No query found against the key PK " + Global.Dictionary["PK"]);
                    //Close connection and return false;
                    conn.Close();
                    return null;
                }
                else if (iRowCount > 1)
                {
                    Global.fUpdateExecutionLog(LogType.error, "More than one record returned against PK " + Global.Dictionary["PK"]);
                    conn.Close();
                    return null;
                }
                else
                {
                    //Get DB Type
                    strDBType = DT.Rows[0].ItemArray[fGetColumnName(DT, "DB_TYPE")].ToString();

                    //DB Sync Time
                    strDBSyncTime = DT.Rows[0].ItemArray[fGetColumnName(DT, "DB_SYNC_TIME")].ToString();
                    if (strDBSyncTime == "") strDBSyncTime = "1";

                    //Save To KR
                    strSaveToKR = DT.Rows[0].ItemArray[fGetColumnName(DT, "SAVE_TO_KR")].ToString();

                    //Save Param Name
                    strSaveParamName = DT.Rows[0].ItemArray[fGetColumnName(DT, "SAVE_PARAM_NAME")].ToString();

                    //expected result
                    strExpectedResult = DT.Rows[0].ItemArray[fGetColumnName(DT, "EXP_RESULTS")].ToString();

                    //Form query
                    strQuery = "";
                    int j = 1;
                    while (fGetColumnName(DT, "SQL_" + j) != -1 && DT.Rows[0].ItemArray[fGetColumnName(DT, "SQL_" + j)].ToString() != "")
                    {
                        strQuery = strQuery + DT.Rows[0].ItemArray[fGetColumnName(DT, "SQL_" + j)].ToString();
                        j++;
                    }

                    strQuery = strQuery.Trim();

                    //Check for empty
                    if (strQuery == "")
                    {
                        conn.Close();
                        return null;
                    }

                    //Replace all params from SQL
                    foreach (KeyValuePair<string, string> KVP in Global.Dictionary)
                    {
                        //in SQL
                        strQuery = strQuery.Replace("<" + KVP.Key + ">", KVP.Value);

                        //In Expected result
                        if (strExpectedResult != "" && strExpectedResult != null)
                        {
                            strExpectedResult = strExpectedResult.Replace("<" + KVP.Key + ">", KVP.Value);
                        }
                        else
                        {
                            strExpectedResult = "";
                        }
                    }
                }

                //clsong connection
                conn.Close();
                DS.Clear();

                //Replace Keep Refer Params from SQL
                strQuery = fReplaceSpecialParameterInSQL(strQuery);



                //Reconnect to required DB
                objConn = fConnectToDB(strDBType);
                if (objConn == null)
                {
                    Global.fUpdateExecutionLog(LogType.error, "Failed to establish the connection to the DB " + strDBType);
                    return null;
                }

                //initialize sync
                int iSync = 0;

                //Loop to execute query
                while (iSync <= Convert.ToInt32(strDBSyncTime) && ((resDBCheck == false && ExpectedResFlag == true) || iSync == 0))
                {
                    for (int iLoop = 1; iLoop <= Convert.ToInt32(strDBSyncTime); iLoop++)
                    {
                        //Check is its a select Query
                        if (strQuery.ToUpper().Contains("SELECT"))
                        {
                            DS = fExecuteSelectQuery(strQuery, objConn);
                            if (DS == null) continue;

                            //set flag
                            Flag = true;
                            break;
                        }
                        else
                        {
                            if (fExecuteInsertUpdateQuery(strQuery, objConn) == false)
                            {
                                Flag = false;
                                break;
                            }
                        }
                    }

                    //Validation of executed query
                    if (Flag == false)
                    {
                        Global.fUpdateExecutionLog(LogType.error, "Failed to execute query " + strQuery);
                        objConn.Close();
                        return null;
                    }

                    //Get RowCount
                    DT = DS.Tables[0];
                    int iRecRows = DT.Rows.Count;
                    int iColCnt = DT.Columns.Count;

                    //Check for NUM_SQL_RESULTS
                    if (Global.Dictionary.ContainsKey("NUM_SQL_RESULTS") == false) Global.Dictionary.Add("NUM_SQL_RESULTS", "");

                    if (Global.Dictionary["NUM_SQL_RESULTS"].Trim() == "")
                    {
                        //Loop through columns
                        for (int i = 0; i < iColCnt; i++)
                        {
                            outDictionary.Add(DT.Columns[i].ColumnName, DT.Rows[0].ItemArray[i].ToString());
                        }
                    }
                    else if (Global.Dictionary["NUM_SQL_RESULTS"].Trim() != "")
                    {
                        //loop through No of SQL results
                        for (int j = 0; j < Convert.ToInt32(Global.Dictionary["NUM_SQL_RESULTS"]); j++)
                        {
                            //Loop through columns
                            for (int i = 0; i < iColCnt; i++)
                            {
                                if (j == 0) outDictionary.Add(DT.Columns[i].ColumnName, DT.Rows[j].ItemArray[i].ToString());
                                else outDictionary[DT.Columns[i].ColumnName] = outDictionary[DT.Columns[i].ColumnName] + "," + DT.Rows[j].ItemArray[i].ToString();
                            }
                        }
                    }


                    //Closing connection
                    objConn.Close();

                    //Declare array to fetch expected result
                    //string[] arrExpRes;
                    char[] delimiters = { ';' };

                    //Check if expected result is populated
                    if (strExpectedResult.Trim() != "")
                    {
                        ExpectedResFlag = true;
                        //arrExpRes = strExpectedResult.Split(delimiters);
                    }
                    else
                    {
                        ExpectedResFlag = false;
                    }

                    //if Saveto KR is true save details to keep refer with set reference verification data
                    if (strSaveToKR.ToUpper().Trim() == "TRUE")
                    {
                        //Set SQL in DBSQL param in dict
                        if (Global.Dictionary.ContainsKey("DBSQL"))
                        {
                            Global.Dictionary["DBSQL"] = strQuery;
                        }
                        else
                        {
                            Global.Dictionary.Add("DBSQL", strQuery);
                        }

                        //Save values in KR
                        if (fSetReferenceVerificationData(strSaveParamName, outDictionary) == false) return null;

                    }


                    //Compare expected results
                    if (ExpectedResFlag == true)
                    {
                        //Create Fail test Key in DIctionary
                        if (Global.Dictionary.ContainsKey("FAIL_TEST"))
                        {
                            Global.Dictionary["FAIL_TEST"] = "N";
                        }
                        else
                        {
                            Global.Dictionary.Add("FAIL_TEST", "N");
                        }

                        //Call the function fDBCompareArray and store the result to the resDBCheck flag
                        resDBCheck = fDBCompareArray(strExpectedResult.Split(delimiters), outDictionary);

                        if (resDBCheck == false)
                        {
                            outDictionary = null;
                            outDictionary = new Dictionary<string, string>();
                        }

                        Global.Dictionary["FAIL_TEST"] = "";
                    }
                    else
                    {
                        resDBCheck = true;
                    }


                    //Increment iSync
                    iSync++;
                }
                return outDictionary;

            }
            catch (Exception e)
            {
                Global.fUpdateExecutionLog(LogType.error, "Got Exception while executing the fDBCheck function. Exception is " + e);
                conn.Close();
                objConn.Close();
                return null;
            }

        }

        //*****************************************************************************************
        //*	Name		    : fGetColumnName
        //*	Description	    : FUnction to Column names of record
        //*	Author		    : Aniket Gadre
        //*	Input Params	: DataTable, ColumnName
        //*	Return Values	: int : No of columns
        //*****************************************************************************************
        public int fGetColumnName(DataTable DT, string strColumnName)
        {
            int iColCnt = DT.Columns.Count;

            for (int z = 0; z < iColCnt; z++)
            {
                if (DT.Columns[z].ColumnName.Trim() == strColumnName.Trim()) return z;
            }

            //return 
            return -1;
        }

        //*****************************************************************************************
        //*	Name		    : fDBCompareArray
        //*	Description	    : Compare two array and report the result
        //*	Author		    : Aniket Gadre
        //*	Input Params	: String [] arrExp - Array of expected result
        //*					  Dictionary outDictTemp - Dict containing the value fetched from the DB	
        //*	Return Values	: boolean - Based on Success or failure
        //*****************************************************************************************
        public bool fDBCompareArray(string[] arrExpRes, Dictionary<string, string> octDictTemp)
        {
            //log
            Global.fUpdateExecutionLog(LogType.debug, "Executing function : fDBCompareArray");

            //Declare
            string strExpectedValue, strActualValue;
            bool flag = false;
            string[] arrParamValue;

            //get values from dict into array
            try
            {
                arrParamValue = octDictTemp.Values.ToArray();
            }
            catch (Exception e)
            {
                //log
                Global.fUpdateExecutionLog(LogType.error, "Converting dictionary values into Array Failed. Exception " + e + " occured");
                return false;
            }

            //check array size
            if (arrExpRes.Count() != arrParamValue.Count())
            {
                //log & Report
                Global.fUpdateExecutionLog(LogType.error, "Size of Expcted Result array and Dictionary values array did not match");
                return false;
            }

            //Loop through arrays and save values in KEEP_REFER
            for (int iLoop = 0; iLoop < arrParamValue.Count(); iLoop++)
            {

                //Temp
                string strTemp = "";
                char[] charDel = { '>', '<', '=', '~' };
                string[] strDel = { ">=", "<=", "==", "<>" };

                //Fetch expected and actual value
                strExpectedValue = arrExpRes[iLoop].Trim();
                strActualValue = arrParamValue[iLoop].Trim();

                if (strExpectedValue.Contains(">="))
                {
                    strTemp = strExpectedValue.Split(strDel, StringSplitOptions.None)[1];
                    if (fValidateComparisonValues(strTemp, strActualValue) == false) return false;
                    flag = (Convert.ToInt64(strActualValue) >= Convert.ToInt64(strTemp));
                }
                else if (strExpectedValue.Contains("<="))
                {
                    strTemp = strExpectedValue.Split(strDel, StringSplitOptions.None)[1];
                    if (fValidateComparisonValues(strTemp, strActualValue) == false) return false;
                    flag = (Convert.ToInt64(strActualValue) <= Convert.ToInt64(strTemp));
                }
                else if (strExpectedValue.Contains("<>"))
                {
                    strTemp = strExpectedValue.Split(strDel, StringSplitOptions.None)[1];
                    //if (fValidateComparisonValues(strTemp, strActualValue) == false) return false;
                    flag = (strActualValue != strTemp);
                }
                else if (strExpectedValue.Contains(">"))
                {
                    strTemp = strExpectedValue.Split(charDel)[1];
                    if (fValidateComparisonValues(strTemp, strActualValue) == false) return false;
                    flag = (Convert.ToInt64(strActualValue) > Convert.ToInt64(strTemp));
                }
                else if (strExpectedValue.Contains("<"))
                {
                    strTemp = strExpectedValue.Split(charDel)[1];
                    if (fValidateComparisonValues(strTemp, strActualValue) == false) return false;
                    flag = (Convert.ToInt64(strActualValue) < Convert.ToInt64(strTemp));
                }
                else if (strExpectedValue.Contains("="))
                {
                    strTemp = strExpectedValue.Split(charDel)[1];
                    //if (fValidateComparisonValues(strTemp, strActualValue) == false) return false;
                    flag = (strActualValue == strTemp);
                }
                else if (strExpectedValue.Contains("~"))
                {
                    strTemp = strExpectedValue.Split(charDel)[1];
                    //if (fValidateComparisonValues(strTemp, strActualValue) == false) return false;
                    flag = strActualValue.Contains(strTemp);
                }
                else
                {
                    //log
                    Global.fUpdateExecutionLog("No operator mentioned in Expected value " + strExpectedValue + ". '=' will be used as default operator");
                    flag = (strExpectedValue == strActualValue);
                }


                //Check Flag Value
                if (flag == false)
                {
                    //log and report
                    Global.fUpdateExecutionLog(LogType.error, "Comparison between Expected and Actual value failed. Expected Value " + strExpectedValue + " and Actual Value " + strActualValue);
                    return false;
                }
                else Global.fUpdateExecutionLog(LogType.info, "Comparison between Expected and Actual value was successfull. Expected Value " + strExpectedValue + " and Actual Value " + strActualValue);

            }

            //return
            return true;
        }

        //*****************************************************************************************
        //*	Name		    : fValidateComparisonValues
        //*	Description	    : Compare two array and report the result
        //*	Author		    : Aniket Gadre
        //*	Input Params	: string 1, string 2
        //*	Return Values	: boolean - Based on Success or failure
        //*****************************************************************************************
        public bool fValidateComparisonValues(string str1, string str2)
        {
            Global.fUpdateExecutionLog(LogType.info, "Entered function: fValidateComparisonValues");
            try
            {
                long int1 = Convert.ToInt64(str1);
                long int2 = Convert.ToInt64(str2);
            }
            catch (Exception e)
            {
                //log
                Global.fUpdateExecutionLog(LogType.error, "Unable to convert strings " + str1 + " and " + str2 + " to long");
                return false;
            }

            //return
            return true;
        }

    }
}
