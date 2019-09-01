#region Help:  Introduction to the script task
/* The Script Task allows you to perform virtually any operation that can be accomplished in
 * a .Net application within the context of an Integration Services control flow. 
 * 
 * Expand the other regions which have "Help" prefixes for examples of specific ways to use
 * Integration Services features within this script task. */
#endregion


#region Namespaces
using System;
using System.IO.Compression;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using SColor = System.Drawing; /// Changing font color
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using Microsoft.CSharp;
using System.Collections; /// Adding Arraylists
using Excel = Microsoft.Office.Interop.Excel; /// Excel Applications
#endregion

namespace ExcelPOC
{
    /// <summary>
    /// ScriptMain is the entry point class of the script.  Do not change the name, attributes,
    /// or parent of this class.
    /// </summary>
	[Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
	public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
	{
        private int i;
        #region Help:  Using Integration Services variables and parameters in a script
        /* To use a variable in this script, first ensure that the variable has been added to 
         * either the list contained in the ReadOnlyVariables property or the list contained in 
         * the ReadWriteVariables property of this script task, according to whether or not your
         * code needs to write to the variable.  To add the variable, save this script, close this instance of
         * Visual Studio, and update the ReadOnlyVariables and 
         * ReadWriteVariables properties in the Script Transformation Editor window.
         * To use a parameter in this script, follow the same steps. Parameters are always read-only.
         * 
         * Example of reading from a variable:
         *  DateTime startTime = (DateTime) Dts.Variables["System::StartTime"].Value;
         * 
         * Example of writing to a variable:
         *  Dts.Variables["User::myStringVariable"].Value = "new value";
         * 
         * Example of reading from a package parameter:
         *  int batchId = (int) Dts.Variables["$Package::batchId"].Value;
         *  
         * Example of reading from a project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].Value;
         * 
         * Example of reading from a sensitive project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].GetSensitiveValue();
         * */

        #endregion

        #region Help:  Firing Integration Services events from a script
        /* This script task can fire events for logging purposes.
         * 
         * Example of firing an error event:
         *  Dts.Events.FireError(18, "Process Values", "Bad value", "", 0);
         * 
         * Example of firing an information event:
         *  Dts.Events.FireInformation(3, "Process Values", "Processing has started", "", 0, ref fireAgain)
         * 
         * Example of firing a warning event:
         *  Dts.Events.FireWarning(14, "Process Values", "No values received for input", "", 0);
         * */
        #endregion

        #region Help:  Using Integration Services connection managers in a script
        /* Some types of connection managers can be used in this script task.  See the topic 
         * "Working with Connection Managers Programatically" for details.
         * 
         * Example of using an ADO.Net connection manager:
         *  object rawConnection = Dts.Connections["Sales DB"].AcquireConnection(Dts.Transaction);
         *  SqlConnection myADONETConnection = (SqlConnection)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Sales DB"].ReleaseConnection(rawConnection);
         *
         * Example of using a File connection manager
         *  object rawConnection = Dts.Connections["Prices.zip"].AcquireConnection(Dts.Transaction);
         *  string filePath = (string)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Prices.zip"].ReleaseConnection(rawConnection);
         * */
        #endregion

        
		/// <summary>
        /// This method is called when this script task executes in the control flow.
        /// Before returning from this method, set the value of Dts.TaskResult to indicate success or failure.
        /// To open Help, press F1.
        /// </summary>
        /// 

		public void Main()
		{
            string SMTPServer = "015-smtp-out.aviva.com";
            string MailFrom = "yrrpts1@aviva.com";
            string MailTo = "babulal.ram@aviva.com";
            string MailCC = "babulal.ram@aviva.com";

            string QueryTruncTables = Dts.Variables["User::SQLQuery_TruncateTables"].Value.ToString();
            string BOOutputFilewithPath = Dts.Variables["User::FILE_BOOutputFilewithPath"].Value.ToString(); 
            string Database = Dts.Variables["User::Database"].Value.ToString();
            string Server = Dts.Variables["User::Server"].Value.ToString();

            string SQLConnString = @"Data Source=" + Server + ";Initial Catalog=" + Database + ";Integrated Security=SSPI;";

            TruncateSQLTables(SQLConnString, BOOutputFilewithPath, QueryTruncTables, MailFrom, MailTo, MailCC, SMTPServer);

            Notification_To_SQL(SQLConnString, BOOutputFilewithPath, MailFrom, MailTo, MailCC, SMTPServer);
            Occurrences_To_SQL(SQLConnString, BOOutputFilewithPath, MailFrom, MailTo, MailCC, SMTPServer);
            Settlements_To_SQL(SQLConnString, BOOutputFilewithPath, MailFrom, MailTo, MailCC, SMTPServer);
            Outstanding_To_SQL(SQLConnString, BOOutputFilewithPath, MailFrom, MailTo, MailCC, SMTPServer);

            Dts.TaskResult = (int)ScriptResults.Success;
        }

        private void TruncateSQLTables(string SQLConnString, string BOOutputFilewithPath, string QueryTruncTables, string MailFrom, string MailTo, string MailCC, string SMTPServer)
        {           
                try
                {
                    SqlConnection SQLCON = new SqlConnection(SQLConnString);
                    SQLCON.Open();
                    SqlCommand SQLCommand = new SqlCommand(QueryTruncTables, SQLCON);
                    SQLCommand.ExecuteScalar();
                    SQLCON.Close();
                    
                }
                catch (Exception ex)
                {
                    string MailBody, MailSub = "SP Team - Tesco Weekly - Truncate of tables failed";
                    MailBody = "Hello Team </BR></BR>";
                    MailBody += string.Format("Truncate of tables failed due to:</BR><b>Error message:</b>{1}", ex.Message);

                    SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);

                    Dts.TaskResult = (int)ScriptResults.Failure;
                }
            
        }

        private void Notification_To_SQL(string SQLConnString, string BOOutputFilewithPath, string MailFrom, string MailTo, string MailCC, string SMTPServer)
        {
            string SheetName = "Notifications";
            string TableName = "AUT.tbl_FACT_Notifications";
            string ExceltoLoad = BOOutputFilewithPath;

            if (File.Exists(ExceltoLoad))
            {
                String ExcelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", ExceltoLoad);
                try
                {
                    //Create Connection to Excel work book 
                    using (OleDbConnection excelConnection = new OleDbConnection(ExcelConnString))
                    {
                        //Create OleDbCommand to fetch data from Excel 
                        using (OleDbCommand cmd = new OleDbCommand(@"Select [Cover Level 1 Name], [Claim Type Name], [Product Name], [Cause Code Name], [Week_Notified], 
                                            [CountOfCustomer Claim Number], [SumOfSumOfTotal_Paid], [SumOfSumOfTotal_Estimate], 
                                            [SumOfSumOfTotal_Incurred], [SumOfSumOfAD_Paid], [SumOfSumOfAD_Estimate], [SumOfSumOfTP_Paid], [SumOfSumOfTP_Estimate], 
                                            [SumOfSumOfBI_Paid], [SumOfSumOfBI_Estimate],[SumOfSumOfRec_Paid], [SumOfSumOfRec_Estimate], [Responsibilty Percentage]
                                      FROM [" + SheetName + "$]", excelConnection))
                        {
                            excelConnection.Open();
                            using (OleDbDataReader dReader = cmd.ExecuteReader())
                            {
                                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(SQLConnString))
                                {
                                    sqlBulk.BulkCopyTimeout = 0;
                                    sqlBulk.DestinationTableName = TableName; //Give your Destination table name 
                                    sqlBulk.ColumnMappings.Add("[Cover Level 1 Name]","[Cover Level 1 Name]");
                                    sqlBulk.ColumnMappings.Add("[Claim Type Name]","[Claim Type Name]");
                                    sqlBulk.ColumnMappings.Add("[Product Name]","[Product Name]");
                                    sqlBulk.ColumnMappings.Add("[Cause Code Name]","[Cause Code Name]" );
                                    sqlBulk.ColumnMappings.Add("[Week_Notified]","[Week_Notified]");
                                    sqlBulk.ColumnMappings.Add("[CountOfCustomer Claim Number]","[CountOfCustomer Claim Number]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Paid]","[SumOfSumOfTotal_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Estimate]","[SumOfSumOfTotal_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Incurred]","[SumOfSumOfTotal_Incurred]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfAD_Paid]","[SumOfSumOfAD_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfAD_Estimate]","[SumOfSumOfAD_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTP_Paid]","[SumOfSumOfTP_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTP_Estimate]","[SumOfSumOfTP_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfBI_Paid]","[SumOfSumOfBI_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfBI_Estimate]","[SumOfSumOfBI_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfRec_Paid]","[SumOfSumOfRec_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfRec_Estimate]","[SumOfSumOfRec_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[Responsibilty Percentage]","[Responsibilty Percentage]");
                                    sqlBulk.WriteToServer(dReader);
                                }
                            }
                            excelConnection.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    string MailBody, MailSub = "SP Team - Tesco Weekly - Notifications component failed";
                    MailBody = "Hello Team </BR></BR>";
                    MailBody += string.Format("Data load for the table <b>{0}</b> failed due to:</BR><b>Error message:</b>{1}", TableName, ex.Message);

                    SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);
                    Dts.TaskResult = (int)ScriptResults.Failure;
                }
            }
        }

        private void Occurrences_To_SQL(string SQLConnString, string BOOutputFilewithPath, string MailFrom, string MailTo, string MailCC, string SMTPServer)
        {
            string SheetName = "Occurences";
            string TableName = "AUT.tbl_FACT_Occurences";
            string ExceltoLoad = BOOutputFilewithPath;

            if (File.Exists(ExceltoLoad))
            {
                String ExcelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", ExceltoLoad);
                try
                {
                    //Create Connection to Excel work book 
                    using (OleDbConnection excelConnection = new OleDbConnection(ExcelConnString))
                    {
                        //Create OleDbCommand to fetch data from Excel 
                        using (OleDbCommand cmd = new OleDbCommand(@"Select [Cover Level 1 Name], [Claim Type Name], [Product Name], [Cause Code Name], [Week_Occurred], 
                                    [CountOfCustomer Claim Number], [SumOfSumOfTotal_Paid], [SumOfSumOfTotal_Estimate], 
                                    [SumOfSumOfTotal_Incurred], [SumOfSumOfAD_Paid], [SumOfSumOfAD_Estimate], [SumOfSumOfTP_Paid], [SumOfSumOfTP_Estimate], [SumOfSumOfBI_Paid], 
                                    [SumOfSumOfBI_Estimate], [SumOfSumOfRec_Paid], [SumOfSumOfRec_Estimate], [Responsibilty Percentage]
                                    FROM [" + SheetName + "$]", excelConnection))
                        {
                            excelConnection.Open();
                            using (OleDbDataReader dReader = cmd.ExecuteReader())
                            {
                                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(SQLConnString))
                                {
                                    sqlBulk.BulkCopyTimeout = 0;
                                    sqlBulk.DestinationTableName = TableName; //Give your Destination table name 
                                    sqlBulk.ColumnMappings.Add("[Cover Level 1 Name]", "[Cover Level 1 Name]");
                                    sqlBulk.ColumnMappings.Add("[Claim Type Name]", "[Claim Type Name]");
                                    sqlBulk.ColumnMappings.Add("[Product Name]", "[Product Name]");
                                    sqlBulk.ColumnMappings.Add("[Cause Code Name]", "[Cause Code Name]");
                                    sqlBulk.ColumnMappings.Add("[Week_Occurred]", "[Week_Occurred]");
                                    sqlBulk.ColumnMappings.Add("[CountOfCustomer Claim Number]", "[CountOfCustomer Claim Number]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Paid]", "[SumOfSumOfTotal_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Estimate]", "[SumOfSumOfTotal_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Incurred]", "[SumOfSumOfTotal_Incurred]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfAD_Paid]", "[SumOfSumOfAD_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfAD_Estimate]", "[SumOfSumOfAD_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTP_Paid]", "[SumOfSumOfTP_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTP_Estimate]", "[SumOfSumOfTP_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfBI_Paid]", "[SumOfSumOfBI_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfBI_Estimate]", "[SumOfSumOfBI_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfRec_Paid]", "[SumOfSumOfRec_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfRec_Estimate]", "[SumOfSumOfRec_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[Responsibilty Percentage]", "[Responsibilty Percentage]");
                                    sqlBulk.WriteToServer(dReader);
                                }
                            }
                            excelConnection.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    string MailBody, MailSub = "SP Team - REG1209211354W Tesco Weekly - Occurences component failed";
                    MailBody = "Hello Team </BR></BR>";
                    MailBody += string.Format("Data load for the table <b>{0}</b> failed due to:</BR><b>Error message:</b>{1}", TableName, ex.Message);

                    SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);
                    Dts.TaskResult = (int)ScriptResults.Failure;
                }
            }       
        }

        private void Settlements_To_SQL(string SQLConnString, string BOOutputFilewithPath, string MailFrom, string MailTo, string MailCC, string SMTPServer)
        {
            string SheetName = "Settled";
            string TableName = "AUT.tbl_Fact_Settled";
            string ExceltoLoad = BOOutputFilewithPath;

            if (File.Exists(ExceltoLoad))
            {
                String ExcelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", ExceltoLoad);
                try
                {
                    //Create Connection to Excel work book 
                    using (OleDbConnection excelConnection = new OleDbConnection(ExcelConnString))
                    {
                        //Create OleDbCommand to fetch data from Excel 
                        using (OleDbCommand cmd = new OleDbCommand(@"Select [Cover Level 1 Name], [Claim Type Name], [Product Name], [Cause Code Name], [Week_Settled], [Nil_Claim], 
                                        [Outcome_Type], [Claim Indicator Name], [Lifecycle], [CountOfCustomer Claim Number], [SumOfSumOfTotal_Paid], 
                                        [SumOfSumOfAD_Paid], [SumOfSumOfTP_Paid], [SumOfSumOfBI_Paid], [SumOfSumOfRec_Paid], [Responsibilty Percentage]
                                        FROM [" + SheetName + "$]", excelConnection))
                        {
                            excelConnection.Open();
                            using (OleDbDataReader dReader = cmd.ExecuteReader())
                            {
                                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(SQLConnString))
                                {
                                    sqlBulk.BulkCopyTimeout = 0;
                                    sqlBulk.DestinationTableName = TableName; //Give your Destination table name 
                                    sqlBulk.ColumnMappings.Add("[Cover Level 1 Name]", "[Cover Level 1 Name]");
                                    sqlBulk.ColumnMappings.Add("[Claim Type Name]", "[Claim Type Name]");
                                    sqlBulk.ColumnMappings.Add("[Product Name]", "[Product Name]");
                                    sqlBulk.ColumnMappings.Add("[Cause Code Name]", "[Cause Code Name]");
                                    sqlBulk.ColumnMappings.Add("[Week_Settled]", "[Week_Settled]");
                                    sqlBulk.ColumnMappings.Add("[Nil_Claim]", "[Nil_Claim]");
                                    sqlBulk.ColumnMappings.Add("[Outcome_Type]", "[Outcome_Type]");
                                    sqlBulk.ColumnMappings.Add("[Claim Indicator Name]", "[Claim Indicator Name]");
                                    sqlBulk.ColumnMappings.Add("[Lifecycle]", "[Lifecycle]");
                                    sqlBulk.ColumnMappings.Add("[CountOfCustomer Claim Number]", "[CountOfCustomer Claim Number]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Paid]", "[SumOfSumOfTotal_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfAD_Paid]", "[SumOfSumOfAD_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTP_Paid]", "[SumOfSumOfTP_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfBI_Paid]", "[SumOfSumOfBI_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfRec_Paid]", "[SumOfSumOfRec_Paid]");
                                    sqlBulk.ColumnMappings.Add("[Responsibilty Percentage]", "[Responsibilty Percentage]");
                                    sqlBulk.WriteToServer(dReader);
                                }
                            }
                            excelConnection.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    string MailBody, MailSub = "SP Team - REG1209211354W Tesco Weekly - Settled component failed";
                    MailBody = "Hello Team </BR></BR>";
                    MailBody += string.Format("Data load for the table <b>{0}</b> failed due to:</BR><b>Error message:</b>{1}", TableName, ex.Message);

                    SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);
                    Dts.TaskResult = (int)ScriptResults.Failure;
                }
            }        
        }

        private void Outstanding_To_SQL(string SQLConnString, string BOOutputFilewithPath, string MailFrom, string MailTo, string MailCC, string SMTPServer)
        {
            string SheetName = "Outstanding";
            string TableName = "AUT.tbl_FACT_Outstanding";
            string ExceltoLoad = BOOutputFilewithPath;

            if (File.Exists(ExceltoLoad))
            {
                String ExcelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", ExceltoLoad);
                try
                {
                    //Create Connection to Excel work book 
                    using (OleDbConnection excelConnection = new OleDbConnection(ExcelConnString))
                    {
                        //Create OleDbCommand to fetch data from Excel 
                        using (OleDbCommand cmd = new OleDbCommand(@"Select [Cover Level 1 Name], [Claim Type Name], [Cause Code Name], [Product Name], [Age_Banding], 
                                        [Lifecycle], [CountOfCustomer Claim Number], [SumOfSumOfTotal_Paid], [SumOfSumOfTotal_Estimate], [SumOfSumOfTotal_Incurred], 
                                        [SumOfSumOfAD_Paid], [SumOfSumOfAD_Estimate], [SumOfSumOfTP_Paid], [SumOfSumOfTP_Estimate], [SumOfSumOfBI_Paid], 
                                        [SumOfSumOfBI_Estimate], [SumOfSumOfRec_Paid], [SumOfSumOfRec_Estimate], [Responsibilty Percentage]
                                        FROM [" + SheetName + "$]", excelConnection))
                        {
                            excelConnection.Open();
                            using (OleDbDataReader dReader = cmd.ExecuteReader())
                            {
                                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(SQLConnString))
                                {
                                    sqlBulk.BulkCopyTimeout = 0;
                                    sqlBulk.DestinationTableName = TableName; //Give your Destination table name 
                                    sqlBulk.ColumnMappings.Add("[Cover Level 1 Name]", "[Cover Level 1 Name]");
                                    sqlBulk.ColumnMappings.Add("[Claim Type Name]", "[Claim Type Name]");
                                    sqlBulk.ColumnMappings.Add("[Cause Code Name]", "[Cause Code Name]");
                                    sqlBulk.ColumnMappings.Add("[Product Name]", "[Product Name]");
                                    sqlBulk.ColumnMappings.Add("[Age_Banding]", "[Age_Banding]");
                                    sqlBulk.ColumnMappings.Add("[Lifecycle]", "[Lifecycle]");
                                    sqlBulk.ColumnMappings.Add("[CountOfCustomer Claim Number]", "[CountOfCustomer Claim Number]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Paid]", "[SumOfSumOfTotal_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Estimate]", "[SumOfSumOfTotal_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Incurred]", "[SumOfSumOfTotal_Incurred]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfAD_Paid]", "[SumOfSumOfAD_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfAD_Estimate]", "[SumOfSumOfAD_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTP_Paid]", "[SumOfSumOfTP_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfTP_Estimate]", "[SumOfSumOfTP_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfBI_Paid]", "[SumOfSumOfBI_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfBI_Estimate]", "[SumOfSumOfBI_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfRec_Paid]", "[SumOfSumOfRec_Paid]");
                                    sqlBulk.ColumnMappings.Add("[SumOfSumOfRec_Estimate]", "[SumOfSumOfRec_Estimate]");
                                    sqlBulk.ColumnMappings.Add("[Responsibilty Percentage]", "[Responsibilty Percentage]");
                                    sqlBulk.WriteToServer(dReader);
                                }
                            }
                            excelConnection.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    string MailBody, MailSub = "SP Team - REG1209211354W Tesco Weekly - Outstanding component failed";
                    MailBody = "Hello Team </BR></BR>";
                    MailBody += string.Format("Data load for the table <b>{0}</b> failed due to:</BR><b>Error message:</b>{1}", TableName, ex.Message);

                    SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);
                    Dts.TaskResult = (int)ScriptResults.Failure;
                }
            }
        }

        private void SendMail(string MailFrom, string MailTo, string MailCC, string MailSub, string SMTPServer, string MailBody)
        {
            MailMessage htmlMessage = new MailMessage();
            SmtpClient mySmtpClient = new SmtpClient(SMTPServer);
            MailBody += "</BR></BR>";
            MailBody += "Thank you,</BR>";
            MailBody += "YourReports";
            htmlMessage.Body = MailBody.ToString();
            htmlMessage = new MailMessage(MailFrom, MailTo, MailSub, htmlMessage.Body.ToString());
            htmlMessage.IsBodyHtml = true;
            mySmtpClient.Credentials = CredentialCache.DefaultNetworkCredentials;
            mySmtpClient.Send(htmlMessage);
        }

        #region ScriptResults declaration
        /// <summary>
        /// This enum provides a convenient shorthand within the scope of this class for setting the
        /// result of the script.
        /// 
        /// This code was generated automatically.
        /// </summary>
        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        };
        #endregion

	}
}