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
using System.Data.OleDb;
using System.Data.SqlClient;
using SColor = System.Drawing; /// Changing font color
using Microsoft.SqlServer.Dts.Runtime;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections;
using System.Windows.Forms;
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
            //string ExcelFilePath = @"C:\VSS_CHECK_OUT\MI Automation\C practices\Refresh\REG0701021459M - Monthly Flash Variance Report June 2018.xlsx";
            //string Password = Dts.Variables["User::Password"].Value.ToString();
            //string ReportingWeek = null;

            //Excel.Application ExcelApp = new Excel.Application();               // Initialize Excel Application
            //ExcelApp.DisplayAlerts = false;

            //Excel.Workbook WB = ExcelApp.Workbooks.Open(ExcelFilePath);              // Initialize Excel Workbook
            //Excel.Worksheets WS = ExcelApp.ActiveSheet as Excel.Worksheets;   // Initialize Excel Worksheet -- Not Required

            //string Cells = "U46:AD60";
            //ArrayList toHide = new ArrayList();
            //ArrayList toDelete = new ArrayList();
            //ArrayList toProtect = new ArrayList();

            ////Assigining toDelete data from dts.Variable to Arraylist (toDelete)
            //foreach (string data in Dts.Variables["User::toDelete"].Value.ToString().Split(','))
            //{
            //    toDelete.Add(data);
            //}

            ////Assigining toHide data from dts.Variable to Arraylist (toHide)
            //foreach (string data in Dts.Variables["User::toHide"].Value.ToString().Split(','))
            //{
            //    toHide.Add(data);
            //}

            ////Assigining toProtect data from dts.Variable to Arraylist (toProtect)
            //foreach (string data in Dts.Variables["User::toProtect"].Value.ToString().Split(','))
            //{
            //    toProtect.Add(data);
            //}
           
            try
            {
                //DownloadAttachment();
                //ExportDatatoExcel(ExcelApp, WB);                    // Export data from SQL Server to Excel
                //InsertRow(ExcelApp, WB);
                //RefreshSheets(ExcelApp ,WB);                        // Refresh Excel sheets
                //ChangeCellColor(WB, Cells);                         // Change cellcolor of cells
                //HideSheet(WB, toHide);                              // Hide Sheet                       
                //FormatSheetData(WB);                                // Format style/fonts
                //UpdateFrontSheetData(WB, ReportingWeek);            // Update latest date in front sheet 
                //ProtectSheet(ExcelApp, WB, toProtect, Password);    // Protect Sheet
                //DeleteSheet(WB, toDelete);                          // Delete sheet
                //DeleteSheet_OldCode(WB, toDelete);                  // Delete Sheet, Not used as the approach was time consuming due to more number of Excel sheet iterations.
                //DeleteColumn(WB);                                   // Delete Column based on index         
                //ApplyFormulaToCells(ExcelApp, WB);                  // DragCells to apply formula for new cells

                Notification_To_SQL();

                //WB.Save();
                //WB.Close(0);
                //ExcelApp.Quit();
                //GC.Collect();
            }

            catch (Exception E)
            {
                throw E;
            }

            finally
            {
                //WB.Save();
                //WB.Close(0);
                //ExcelApp.Quit();
                GC.Collect();
                //Dts.TaskResult = (int)ScriptResults.Failure;
            }           

            Dts.TaskResult = (int)ScriptResults.Success;
		}

        private void Notification_To_SQL()
        {
            string SheetName = "NB Avg Prem";
            string TableName = "QRF.TBL_QRF_DIRECT_PREDICTION";
            string ExceltoLoad = @"C:\VSS_CHECK_OUT\MI Automation\C practices\Refresh\Car Pricing Template JW Version.xlsx";
            string SQLConnString = @"Data Source=Missdbserver02;Initial Catalog=MI_SS_Collections;Integrated Security=SSPI;";
            //if (File.Exists(ExceltoLoad))
            //{
                String ExcelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\"", ExceltoLoad);
                try
                {
                    //Create Connection to Excel work book 
                    using (OleDbConnection excelConnection = new OleDbConnection(ExcelConnString))
                    {
                        //Create OleDbCommand to fetch data from Excel 
                        using (OleDbCommand cmd = new OleDbCommand(@"Select * 
                                      FROM [" + SheetName + "$]", excelConnection))
                      
                        {
                            excelConnection.Open();
                            using (OleDbDataReader dReader = cmd.ExecuteReader())
                            {
                                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(SQLConnString))
                                {
                                    sqlBulk.BulkCopyTimeout = 0;
                                    sqlBulk.DestinationTableName = TableName; //Give your Destination table name 
                                    //sqlBulk.ColumnMappings.Add("[Cover Level 1 Name]", "[Cover Level 1 Name]");
                                    //sqlBulk.ColumnMappings.Add("[Claim Type Name]", "[Claim Type Name]");
                                    //sqlBulk.ColumnMappings.Add("[Product Name]", "[Product Name]");
                                    //sqlBulk.ColumnMappings.Add("[Cause Code Name]", "[Cause Code Name]");
                                    //sqlBulk.ColumnMappings.Add("[Week_Notified]", "[Week_Notified]");
                                    //sqlBulk.ColumnMappings.Add("[CountOfCustomer Claim Number]", "[CountOfCustomer Claim Number]");
                                    //sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Paid]", "[SumOfSumOfTotal_Paid]");
                                    //sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Estimate]", "[SumOfSumOfTotal_Estimate]");
                                    //sqlBulk.ColumnMappings.Add("[SumOfSumOfTotal_Incurred]", "[SumOfSumOfTotal_Incurred]");
                                    //sqlBulk.ColumnMappings.Add("[SumOfSumOfAD_Paid]", "[SumOfSumOfAD_Paid]");
                                    //sqlBulk.ColumnMappings.Add("[SumOfSumOfAD_Estimate]", "[SumOfSumOfAD_Estimate]");
                                    //sqlBulk.ColumnMappings.Add("[SumOfSumOfTP_Paid]", "[SumOfSumOfTP_Paid]");
                                    //sqlBulk.ColumnMappings.Add("[SumOfSumOfTP_Estimate]", "[SumOfSumOfTP_Estimate]");
                                    //sqlBulk.ColumnMappings.Add("[SumOfSumOfBI_Paid]", "[SumOfSumOfBI_Paid]");
                                    //sqlBulk.ColumnMappings.Add("[SumOfSumOfBI_Estimate]", "[SumOfSumOfBI_Estimate]");
                                    //sqlBulk.ColumnMappings.Add("[SumOfSumOfRec_Paid]", "[SumOfSumOfRec_Paid]");
                                    //sqlBulk.ColumnMappings.Add("[SumOfSumOfRec_Estimate]", "[SumOfSumOfRec_Estimate]");
                                    //sqlBulk.ColumnMappings.Add("[Responsibilty Percentage]", "[Responsibilty Percentage]");
                                    sqlBulk.WriteToServer(dReader);
                                }
                            }
                            excelConnection.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                   
                }
            //}
        }

        private void DownloadAttachment()
        {
            try
            {
                string FilePath = @"C:\VSS_CHECK_OUT\MI Automation\Poddows\", FileNamewithPath = null, Subj = "ProddowsFolder Test";              
                Outlook.Application OApp = new Outlook.Application();
                Outlook.NameSpace ONameSpace = OApp.GetNamespace("MAPI");                
                //Outlook.MailItem OMail = null; 
                Outlook.MAPIFolder InboxFolder, ProddowsFolder, ProddowsLoadedFolder = null;
                ArrayList MailCollection = new ArrayList();
                ONameSpace.Logon("avivagroup/dashb",null , false, true);
                InboxFolder = ONameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                ProddowsFolder = InboxFolder.Folders["Proddows"];

                //for (int j = 0; j <= InboxFolder.Folders.Count; j++ )
                //{
                //    MessageBox.Show("Name: " + InboxFolder.Folders.Parent + " FolderPath: " + InboxFolder.FolderPath);
                //}

                ProddowsLoadedFolder = ProddowsFolder.Folders["Downloaded"];
                Outlook.Items OItem = InboxFolder.Items;

                foreach (Outlook.MailItem item in ProddowsFolder.Items)
                {
                    if (item != null && item.Subject == Subj)
                    {
                        MailCollection.Add(item);
                    }
                }

                foreach (Outlook.MailItem OMItem in MailCollection)
                {
                    foreach(Outlook.Attachment OAttactment in OMItem.Attachments)
                    {
                        FileNamewithPath = FilePath + OAttactment.FileName;
                        OAttactment.SaveAsFile(FileNamewithPath);
                        
                        OMItem.Move(ProddowsLoadedFolder);
                    }
                    
                }
                ONameSpace.Logoff();
                OItem = null; 
                MailCollection = null; 
                ONameSpace = null; 
                OApp = null;
            }
            catch (Exception e)
            {

            }
        }

        private void InsertRow(Excel.Application ExcelApp, Excel.Workbook WB)
        {
            Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item("Sheet1")); // Activate specific sheet  
            //Excel.Range InRow = WS.Rows.Insert(2, 5);       

            for (i = 1; i <= 5; i++)
            {
                Excel.Range rng = (Excel.Range)WS.Cells[3,1];
                Excel.Range row = rng.EntireRow;
                row.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
            }
        }

        private void ExportDatatoExcel(Excel.Application ExcelApp, Excel.Workbook WB)
        {

            string query = "SELECT TOP 10 AgencyID,AgencyCode,AgencyFullName,CompanyName,Brokerall,CompanyGroupName FROM DBO.tbl_CR_DIM_AgencyMaster";
            string connectionSql = @"Data Source=MISSDEVDBSERVER;Initial Catalog=MI_SS_IBandUKGI;Integrated Security=SSPI;";
            //string file = @"C:\Automation\SQLtoExcel.xlsx";
            //WB = ExcelApp.Workbooks.Open(file);

            using (SqlConnection connection = new SqlConnection(connectionSql))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();
                // SqlDataReader reader = command.ExecuteReader();
                SqlDataAdapter da = new SqlDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);

                Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item("Sheet1"));
                
                try
                {
                    int row = 1; int col = 1;

                    ////////For writing the column Names

                    foreach (DataColumn column in dt.Columns)
                    {
                        //adding columns
                        WS.Cells[row, col] = column.ColumnName;
                        col++;
                    }

                    //reset column and row variables

                    col = 1;
                    row++;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        //adding data
                        foreach (var cell in dt.Rows[i].ItemArray)
                        {
                            WS.Cells[row, col] = cell;
                            col++;
                        }
                        col = 1;
                        row++;
                    }
                    WS.Columns.AutoFit();
                }
                catch (Exception e)
                {
                    throw e;
                }
            }    
        }

        private void DeleteColumn(Excel.Workbook WB)
        {
            Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item("Data")); // Delete specifc columns from specific sheet
            WS.Columns[1].delete(); // Will delete column based on index number provided.
        }

        private void ApplyFormulaToCells(Excel.Application ExcelApp, Excel.Workbook WB)
        {
            ExcelApp.ScreenUpdating = false;
            string VarFormula = "", GrowthFormula = "";
            int StartCellCount = 21, EndCellCount = StartCellCount + 16 - 1, VarianceCell = 5, GrowthCell = 6;
            
            Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item("Sheet1")); // Activate specific sheet - MONTHLY FLASH         
                       
            WS.get_Range("E" + StartCellCount, "E" + EndCellCount).Formula = "=D" + StartCellCount + "-C" + StartCellCount;
            WS.get_Range("F" + StartCellCount, "F" + EndCellCount).Formula = "=IFERROR(E" + StartCellCount + "/C" + StartCellCount +",0)";

            Excel.Range VarianceTotal = (Excel.Range)WS.Cells[EndCellCount + 1, VarianceCell];
            Excel.Range GrowthTotal = (Excel.Range)WS.Cells[EndCellCount + 1, GrowthCell];

            VarFormula = "=SUM(E" + StartCellCount + ":E" + EndCellCount + ")";
            GrowthFormula = "=SUM(F" + StartCellCount + ":F" + EndCellCount + ")";            

            VarianceTotal.Formula = VarFormula;
            GrowthTotal.Formula = GrowthFormula;            
            
            ExcelApp.ScreenUpdating = true;
        }
        
        private void UpdateFrontSheetData(Excel.Workbook WB, string ReportingWeek)
        {
            Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item("Front Sheet")); // Activate specific sheet
            WS.Cells[25, 8].Value = ReportingWeek;
            //WS.Cells[1, 1].Value = "Week 1";
            //WS.Cells[1, 1].Value = "Week 1";
        }
        
        private void FormatSheetData(Excel.Workbook WB)
        {
            int TotalColumns, TotalRows = 0;
            Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item("Data")); // Activate specific sheet

            WS.Columns.ClearFormats();                                              // Clear cell format                    
            WS.Rows.ClearFormats();                                                 // Clear cell format
            
            TotalColumns = WS.UsedRange.Columns.Count;                              // Count of Columns where data is present
            TotalRows = WS.UsedRange.Rows.Count;                                    // Count of rows where data is present
            WS.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;        // Apply AllBorders around data
            WS.UsedRange.Columns.AutoFit();                                         // Auto Fit column size based on data  
            WS.get_Range("A1","V1").Font.Bold = true;                               // Change first row font as Bold
        }

        private void RefreshSheets(Excel.Application ExcelApp ,Excel.Workbook WB)
        {

            //foreach (Excel.Worksheet WS in WB.Worksheets)
            //{
            //    foreach (Excel.QueryTable query in WS.QueryTables)
            //    {
            //        query.BackgroundQuery = false;
            //    }

            WB.RefreshAll();
            //ExcelApp.Application.CalculateUntilAsyncQueriesDone();

            //}
            //WB.Save();
        }

        private void ProtectSheet(Excel.Application ExcelApp, Excel.Workbook WB, ArrayList toProtect, string Password)
        {
            //ExcelApp.Visible = true;
            foreach (string i in toProtect)
            {  
                    ((Excel.Worksheet)WB.Worksheets[i]).Protect(Password);     
            }
            //WB.Save();
            //ExcelApp.Visible = false;
        }    

        private void HideSheet(Excel.Workbook WB, ArrayList toHide)
        {
            foreach (string i in toHide)
            {
                ((Excel.Worksheet)WB.Worksheets[i]).Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            }
            //WB.Save();
        }

        private void DeleteSheet_OldCode(Excel.Workbook WB, ArrayList toDelete)  // Not used as the approach was time consuming due to more number of Excel sheet iterations.
        {
            foreach (Excel.Worksheet Worksheet in WB.Worksheets)
            {
                if (toDelete.Contains(Worksheet.Name))
                {
                    Worksheet.Delete();
                }    
            }
            //WB.Save();
        }

        private void DeleteSheet(Excel.Workbook WB, ArrayList toDelete)
        {
            foreach (string i in toDelete)
            {               
                    Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item(i)); // Activate specific sheet
                    WS.Delete();         
            }
            //WB.Save();
        }

        private void ChangeCellColor(Excel.Workbook WB, string Cells)
        {
            Excel.Worksheet WS = ((Excel.Worksheet)WB.Worksheets.get_Item("Executive Summary"));
            Excel.Range CellRange = WS.get_Range(Cells) as Excel.Range;
            CellRange.Font.Color = SColor.ColorTranslator.ToOle(SColor.Color.White);
            //WB.Save();
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