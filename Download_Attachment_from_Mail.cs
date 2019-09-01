#region Help:  Introduction to the script task
/* The Script Task allows you to perform virtually any operation that can be accomplished in
 * a .Net application within the context of an Integration Services control flow. 
 * 
 * Expand the other regions which have "Help" prefixes for examples of specific ways to use
 * Integration Services features within this script task. */
#endregion


#region Namespaces
using System;
using System.Data;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Collections; /// Adding Arraylists
using Outlook = Microsoft.Office.Interop.Outlook;                          
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading;
#endregion

namespace ST_845cb482744b427b8b9e52156c3408fd
{
    /// <summary>
    /// ScriptMain is the entry point class of the script.  Do not change the name, attributes,
    /// or parent of this class.
    /// </summary>
	[Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
	public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
	{
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
		public void Main()
		{
            string SMTPServer = Dts.Variables["User::SMTP"].Value.ToString();
            string MailFrom = Dts.Variables["User::MailFrom"].Value.ToString();
            string MailTo = Dts.Variables["User::MailTo"].Value.ToString();
            string MailCC = Dts.Variables["User::MailCC"].Value.ToString();
            string MailSub = "";
            bool ErrorStatus = false;

            DownloadAttachment(MailSub, MailFrom, MailTo, MailCC, SMTPServer, ref ErrorStatus);            
            

            if (ErrorStatus == true)
                Dts.TaskResult = (int)ScriptResults.Failure;
            else
                Dts.TaskResult = (int)ScriptResults.Success;

            //ProcessStarted(MailSub, MailFrom, MailTo, MailCC, SMTPServer); 
        }

        private void DownloadAttachment(string MailSub, string MailFrom, string MailTo, string MailCC, string SMTPServer, ref bool ErrorStatus)
        {
            try
            {
                string FilePath = Dts.Variables["User::FOLDER_DownloadAttachmentLocation"].Value.ToString();
                string FileNamewithPath = null;                
                string Expense_Subj = Dts.Variables["User::MAIL_ExpenseSubject"].Value.ToString();
                Outlook.Application OApp = new Outlook.Application();
                Outlook.NameSpace ONameSpace = OApp.GetNamespace("MAPI");
                //Outlook.MailItem OMail = null; 
                Outlook.MAPIFolder InboxFolder, ExpenseFolder, ExpenseLoadedFolder = null;
                ArrayList MailCollection = new ArrayList();
                ONameSpace.Logon("avivagroup/dashb", null, false, true);
                InboxFolder = ONameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                ExpenseFolder = InboxFolder.Folders["Expense"];

                ExpenseLoadedFolder = ExpenseFolder.Folders["Downloaded"];
                Outlook.Items OItem = InboxFolder.Items;

                if (ExpenseFolder.Items.Count == 0)
                {
                    string MailBody;
                    MailSub = "Expense - Couldn't find email in mailbox";
                    MailBody =("<html><div style=font-family:Calibri;font-size:11.0pt;color:#000000>");
                    MailBody += "Hello Team </BR></BR>";
                    MailBody += "There were no Expense email in the DASHB Mailbox to download the required .ZIP attachments</BR></BR>";                    
                    MailBody += "Thank you,</BR>";
                    MailBody += "YourReports";

                    SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);
                    Dts.TaskResult = (int)ScriptResults.Failure;
                }
                else
                {
                    foreach (Outlook.MailItem item in ExpenseFolder.Items)
                    {
                        if (item != null && (item.Subject.Contains(Expense_Subj)))
                        {
                            MailCollection.Add(item);                         
                        }
                    }

                    foreach (Outlook.MailItem OMItem in MailCollection)
                    {
                        foreach (Outlook.Attachment OAttactment in OMItem.Attachments)
                        {
                            if (OAttactment.FileName.ToString().EndsWith(".zip"))
                            {
                                FileNamewithPath = FilePath + OAttactment.FileName;
                                OAttactment.SaveAsFile(FileNamewithPath);
                            }
                        }
                        OMItem.Move(ExpenseLoadedFolder);
                    }
                    ONameSpace.Logoff();
                    OItem = null;
                    MailCollection = null;
                    ONameSpace = null;
                    OApp = null;                    
                }
            }
            catch (Exception ex)
            {
                string MailBody;
                MailSub = "Expense - Download attachment component failed";
                MailBody = "Hello Team </BR></BR>";
                MailBody += string.Format("Download attachment component failed due to:</BR><b>Error message:</b>{0}</BR></BR>", ex.Message);
                MailBody += "Thank you,</BR>";
                MailBody += "YourReports";

                SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);
                UpdateErrorStatus(ref ErrorStatus);                
            }
        }        

        private void ProcessStarted(string MailSub, string MailFrom, string MailTo, string MailCC, string SMTPServer)
        {            
            //MailSub = "Expense - ETL process started";
            //string MailBody = "";
            //SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);
        }

        private void SendMail(string MailFrom, string MailTo, string MailCC, string MailSub, string SMTPServer, string MailBody)
        {
            MailMessage htmlMessage = new MailMessage();
            SmtpClient mySmtpClient = new SmtpClient(SMTPServer);            

            htmlMessage = new MailMessage(MailFrom, MailTo, MailSub, MailBody);
            //htmlMessage.To.Add("babulal.ram@aviva.com");
            htmlMessage.CC.Add(MailCC.ToString());
            htmlMessage.IsBodyHtml = true;
            mySmtpClient.Credentials = CredentialCache.DefaultNetworkCredentials;
            mySmtpClient.Send(htmlMessage);
        }

        private void UpdateErrorStatus(ref bool ErrorStatus)
        {
            ErrorStatus=true;
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