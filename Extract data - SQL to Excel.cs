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

            string Database = Dts.Variables["User::Database"].Value.ToString();
            string Server = Dts.Variables["User::Server"].Value.ToString();

            string SQLConnString = @"Data Source=" + Server + ";Initial Catalog=" + Database + ";Integrated Security=SSPI;"; // SQL Connection string

            Excel.Application ExcelApp = new Excel.Application();               // Initialize Excel Application
            ExcelApp.DisplayAlerts = false;
            string TemplateFilewithPath = Dts.Variables["User::FILE_TemplateFolder"].Value.ToString() + Dts.Variables["User::FILE_FileName"].Value.ToString();

            Excel.Workbook WB = ExcelApp.Workbooks.Open(TemplateFilewithPath);   // Initialize Excel Workbook
            Excel.Worksheets WS = ExcelApp.ActiveSheet as Excel.Worksheets;   // Initialize Excel Worksheet

            CallManagerData_To_Excel(ExcelApp, WB, MailFrom, MailTo, MailCC, SMTPServer);   // Write CallManager data to Template
          
            string FinalOutputFileName = CreateFileFormTemplate();                          // Create a copy of from Template file to Output file          

            WB = ExcelApp.Workbooks.Open(FinalOutputFileName);                              // Initialize Excel Workbook with output file
            WS = ExcelApp.ActiveSheet as Excel.Worksheets;                                  // Initialize Excel Worksheet with output file
                           
            Notifications_To_Excel(ExcelApp, WB, SQLConnString, MailFrom, MailTo, MailCC, SMTPServer);            // Export data from SQL Server to Excel
            Occurences_To_Excel(ExcelApp, WB, SQLConnString, MailFrom, MailTo, MailCC, SMTPServer);               // Export data from SQL Server to Excel
            Settelments_To_Excel(ExcelApp, WB, SQLConnString, MailFrom, MailTo, MailCC, SMTPServer);              // Export data from SQL Server to Excel
            Outstandings_To_Excel(ExcelApp, WB, SQLConnString, MailFrom, MailTo, MailCC, SMTPServer);             // Export data from SQL Server to Excel
            
            WB.Save();
            WB.Close(0);
            ExcelApp.Quit();
            GC.Collect();            

            Dts.TaskResult = (int)ScriptResults.Success;
        }

        private string CreateFileFormTemplate()
        {            
                string SQLConnString = @"Data Source=MissDBServer01;Initial Catalog=Callmanager;Integrated Security=SSPI;";
                string TemplateFolder = Dts.Variables["User::FILE_TemplateFolder"].Value.ToString();
                string OutputFolder = Dts.Variables["User::FILE_OutputFolder"].Value.ToString();
                string FileName = Dts.Variables["User::FILE_FileName"].Value.ToString();
                string QueryWeekName = Dts.Variables["User::QueryWeekName"].Value.ToString();
                string WeekName = "";

                // Generating WeekNumber 
                SqlConnection SQLCON = new SqlConnection(SQLConnString);
                SQLCON.Open();
                SqlCommand SQLCommand = new SqlCommand(QueryWeekName, SQLCON);
                WeekName = (string)SQLCommand.ExecuteScalar();
                SQLCON.Close();

                Dts.Variables["User::FILE_OutputFolderwithFileName"].Value = OutputFolder + FileName + "-" + WeekName + ".xls";
                Dts.Variables["User::FILE_ZipFileNamewithPath"].Value = OutputFolder + FileName + "-" + WeekName + ".zip";
                Dts.Variables["User::FILE_ZipFileName"].Value = FileName + "-" + WeekName + ".zip";
                Dts.Variables["User::WeekNumber"].Value = (Convert.ToInt16(WeekName.Substring(WeekName.Length - 2, 2)));

                //MessageBox.Show(Dts.Variables["User::WeekNumber"].Value.ToString());

                string FinalOutputFileNamewithPath = OutputFolder + FileName + "-" + WeekName + ".xls";

                string TemplateFile = System.IO.Path.Combine(TemplateFolder, FileName + ".xls");
                if (!System.IO.Directory.Exists(OutputFolder))
                {
                    System.IO.Directory.CreateDirectory(OutputFolder);
                }
                System.IO.File.Copy(TemplateFile, FinalOutputFileNamewithPath, true);

                return FinalOutputFileNamewithPath;
          

        }

        private void Notifications_To_Excel(Excel.Application ExcelApp, Excel.Workbook WB, string SQLConnString, string MailFrom, string MailTo, string MailCC, string SMTPServer)                                 
        {
            try
            {

            string query = "SELECT [Cover Level 1 Name],[Claim Type Name],[Product Name],[Cause Code Name]" +
                           ",[Week_Notified],[CountOfCustomer Claim Number],[SumOfSumOfTotal_Paid],[SumOfSumOfTotal_Estimate]" +
                           ",[SumOfSumOfTotal_Incurred],[SumOfSumOfAD_Paid],[SumOfSumOfAD_Estimate],[SumOfSumOfTP_Paid]" +
                           ",[SumOfSumOfTP_Estimate],[SumOfSumOfBI_Paid],[SumOfSumOfBI_Estimate],[SumOfSumOfRec_Paid]" +
                           ",[SumOfSumOfRec_Estimate],[Responsibilty Percentage] FROM [AUT].[tbl_FACT_Notifications]";

            using (SqlConnection connection = new SqlConnection(SQLConnString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();
                SqlDataAdapter da = new SqlDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);

                Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item("Notifications_Dataset"));

                    int firstRow = 2; int firstCol = 1; int lastRow = 0; int lastCol = 0;
                    lastRow = dt.Rows.Count;
                    lastCol = dt.Columns.Count;
                    Excel.Range top = WS.Cells[firstRow, firstCol];
                    Excel.Range bottom = WS.Cells[lastRow, lastCol];
                    Excel.Range all = (Excel.Range)WS.get_Range(top, bottom);
                    object[,] arrayDT = new object[dt.Rows.Count, dt.Columns.Count];
                   
                    for (int i = 0; i < dt.Rows.Count; i++)
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            arrayDT[i, j] = dt.Rows[i][j];
                        }
                    all.Value2 = arrayDT;

                    WS.Columns.AutoFit();

                }
            }
                catch (Exception ex)
                {
                    string MailBody, MailSub = "SP Team - Tesco Weekly - Notifications data from SQL to Excel failed";
                    MailBody = "Hello Team </BR></BR>";
                    MailBody += string.Format("Data load for Notifications from SQL to Excel failed due to:</BR><b>Error message:</b>{0}", ex.Message);

                    SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);

                    WB.Save();
                    WB.Close(0);
                    ExcelApp.Quit();
                    GC.Collect();   

                    Dts.TaskResult = (int)ScriptResults.Failure;
                }
            
        }

        private void Occurences_To_Excel(Excel.Application ExcelApp, Excel.Workbook WB, string SQLConnString, string MailFrom, string MailTo, string MailCC, string SMTPServer)
        {
         try
         {

            string query = "SELECT [Cover Level 1 Name],[Claim Type Name],[Product Name],[Cause Code Name]"+
                            ",[Week_Occurred],[CountOfCustomer Claim Number],[SumOfSumOfTotal_Paid]"+
                            ",[SumOfSumOfTotal_Estimate],[SumOfSumOfTotal_Incurred],[SumOfSumOfAD_Paid]"+
                            ",[SumOfSumOfAD_Estimate],[SumOfSumOfTP_Paid],[SumOfSumOfTP_Estimate]"+
                            ",[SumOfSumOfBI_Paid],[SumOfSumOfBI_Estimate],[SumOfSumOfRec_Paid]"+
                            ",[SumOfSumOfRec_Estimate],[Responsibilty Percentage] FROM [AUT].[tbl_Fact_Occurences]";

            using (SqlConnection connection = new SqlConnection(SQLConnString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();
                SqlDataAdapter da = new SqlDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);

                Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item("Occurrences_Dataset"));

                
                    int firstRow = 2; int firstCol = 1; int lastRow = 0; int lastCol = 0;
                    lastRow = dt.Rows.Count;
                    lastCol = dt.Columns.Count;
                    Excel.Range top = WS.Cells[firstRow, firstCol];
                    Excel.Range bottom = WS.Cells[lastRow, lastCol];
                    Excel.Range all = (Excel.Range)WS.get_Range(top, bottom);
                    object[,] arrayDT = new object[dt.Rows.Count, dt.Columns.Count];

                    for (int i = 0; i < dt.Rows.Count; i++)
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            arrayDT[i, j] = dt.Rows[i][j];
                        }
                    all.Value2 = arrayDT;

                    WS.Columns.AutoFit();
                }
         }
                catch (Exception ex)
                {
                    string MailBody, MailSub = "SP Team - Tesco Weekly - Occurences data from SQL to Excel failed";
                    MailBody = "Hello Team </BR></BR>";
                    MailBody += string.Format("Data load for Occurences from SQL to Excel failed due to:</BR><b>Error message:</b>{0}", ex.Message);

                    SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);

                    WB.Save();
                    WB.Close(0);
                    ExcelApp.Quit();
                    GC.Collect();

                    Dts.TaskResult = (int)ScriptResults.Failure;
                }
            
        }

        private void Settelments_To_Excel(Excel.Application ExcelApp, Excel.Workbook WB, string SQLConnString, string MailFrom, string MailTo, string MailCC, string SMTPServer)
        {
            try
            {

            string query = "SELECT [Cover Level 1 Name],[Claim Type Name],[Product Name],[Cause Code Name]"+
                            ",[Week_Settled],[Nil_Claim],[Outcome_Type],[Claim Indicator Name]"+
                            ",[Lifecycle],[CountOfCustomer Claim Number],[SumOfSumOfTotal_Paid]"+
                            ",[SumOfSumOfAD_Paid],[SumOfSumOfTP_Paid],[SumOfSumOfBI_Paid]"+
                            ",[SumOfSumOfRec_Paid],[Responsibilty Percentage] FROM [MI_SS_Collections].[AUT].[tbl_Fact_Settled]";

            using (SqlConnection connection = new SqlConnection(SQLConnString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();
                SqlDataAdapter da = new SqlDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);

                Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item("Settlements_Dataset"));

                    int firstRow = 2; int firstCol = 1; int lastRow = 0; int lastCol = 0;
                    lastRow = dt.Rows.Count;
                    lastCol = dt.Columns.Count;
                    Excel.Range top = WS.Cells[firstRow, firstCol];
                    Excel.Range bottom = WS.Cells[lastRow, lastCol];
                    Excel.Range all = (Excel.Range)WS.get_Range(top, bottom);
                    object[,] arrayDT = new object[dt.Rows.Count, dt.Columns.Count];

                    for (int i = 0; i < dt.Rows.Count; i++)
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            arrayDT[i, j] = dt.Rows[i][j];
                        }
                    all.Value2 = arrayDT;

                    WS.Columns.AutoFit();
                }
            }
                catch (Exception ex)
                {
                    string MailBody, MailSub = "SP Team - Tesco Weekly - Settelments data from SQL to Excel failed";
                    MailBody = "Hello Team </BR></BR>";
                    MailBody += string.Format("Data load for Settelments from SQL to Excel failed due to:</BR><b>Error message:</b>{0}", ex.Message);

                    SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);

                    WB.Save();
                    WB.Close(0);
                    ExcelApp.Quit();
                    GC.Collect();

                    Dts.TaskResult = (int)ScriptResults.Failure;
                }
        }

        private void Outstandings_To_Excel(Excel.Application ExcelApp, Excel.Workbook WB, string SQLConnString, string MailFrom, string MailTo, string MailCC, string SMTPServer)
        {
            try
            {

            string query = "SELECT [Cover Level 1 Name],[Claim Type Name],[Cause Code Name],[Product Name]"+
                            ",[Age_Banding],[Lifecycle],[CountOfCustomer Claim Number],[SumOfSumOfTotal_Paid]"+
                            ",[SumOfSumOfTotal_Estimate],[SumOfSumOfTotal_Incurred],[SumOfSumOfAD_Paid]"+
                            ",[SumOfSumOfAD_Estimate],[SumOfSumOfTP_Paid],[SumOfSumOfTP_Estimate]"+
                            ",[SumOfSumOfBI_Paid],[SumOfSumOfBI_Estimate],[SumOfSumOfRec_Paid],[SumOfSumOfRec_Estimate]"+
                            ",[Responsibilty Percentage] FROM [MI_SS_Collections].[AUT].[tbl_Fact_Outstanding]";

            using (SqlConnection connection = new SqlConnection(SQLConnString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();
                SqlDataAdapter da = new SqlDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);

                Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item("Outstanding_Dataset"));

                    int firstRow = 2; int firstCol = 1; int lastRow = 0; int lastCol = 0;
                    lastRow = dt.Rows.Count;
                    lastCol = dt.Columns.Count;
                    Excel.Range top = WS.Cells[firstRow, firstCol];
                    Excel.Range bottom = WS.Cells[lastRow, lastCol];
                    Excel.Range all = (Excel.Range)WS.get_Range(top, bottom);
                    object[,] arrayDT = new object[dt.Rows.Count, dt.Columns.Count];

                    for (int i = 0; i < dt.Rows.Count; i++)
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            arrayDT[i, j] = dt.Rows[i][j];
                        }
                    all.Value2 = arrayDT;

                    WS.Columns.AutoFit();
                }
            }
                catch (Exception ex)
                {
                    string MailBody, MailSub = "SP Team - Tesco Weekly - Outstanding data from SQL to Excel failed";
                    MailBody = "Hello Team </BR></BR>";
                    MailBody += string.Format("Data load for Outstanding from SQL to Excel failed due to:</BR><b>Error message:</b>{0}", ex.Message);

                    SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);

                    WB.Save();
                    WB.Close(0);
                    ExcelApp.Quit();
                    GC.Collect();

                    Dts.TaskResult = (int)ScriptResults.Failure;
                }
        }

        private void CallManagerData_To_Excel(Excel.Application ExcelApp, Excel.Workbook WB, string MailFrom, string MailTo, string MailCC, string SMTPServer)
        {
         try
         {
            string SQLConnString = @"Data Source=MissDBServer01;Initial Catalog=Callmanager;Integrated Security=SSPI;";
            string query = Dts.Variables["User::SQLQuery_Tesco"].Value.ToString();

            using (SqlConnection connection = new SqlConnection(SQLConnString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();
                SqlDataAdapter da = new SqlDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);

                Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item("Data"));

                    WS.Columns.ClearFormats();                                              // Clear cell format                    
                    WS.Rows.ClearFormats();                                                 // Clear cell format

                    int firstRow = 0; int firstCol = 1; int lastRow = 0; int lastCol = 0;
                    firstRow = WS.UsedRange.Rows.Count + 1;
                    lastRow = WS.UsedRange.Rows.Count + dt.Rows.Count;
                    lastCol = dt.Columns.Count;

                    Excel.Range top = WS.Cells[firstRow, firstCol];
                    Excel.Range bottom = WS.Cells[lastRow, lastCol];
                    Excel.Range all = (Excel.Range)WS.get_Range(top, bottom);
                    object[,] arrayDT = new object[dt.Rows.Count, dt.Columns.Count];

                    for (int i = 0; i < dt.Rows.Count; i++)
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            arrayDT[i, j] = dt.Rows[i][j];
                        }
                    all.Value2 = arrayDT;

                    WS.Columns.AutoFit();
                    WB.Save();
                }
            }
                catch (Exception ex)
                {
                    string MailBody, MailSub = "SP Team - Tesco Weekly - CallManager data to Excel failed";
                    MailBody = "Hello Team </BR></BR>";
                    MailBody += string.Format("Data load from CallManager data to Excel failed due to:</BR><b>Error message:</b>{0}", ex.Message);

                    SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);

                    WB.Save();
                    WB.Close(0);
                    ExcelApp.Quit();
                    GC.Collect();

                    Dts.TaskResult = (int)ScriptResults.Failure;
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