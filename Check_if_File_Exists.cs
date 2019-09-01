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
using System.IO;
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
            string SMTPServer = "015-smtp-out.aviva.com";
            string MailFrom = "yrrpts1@aviva.com";
            string MailTo = "babulal.ram@aviva.com";
            string MailCC = "babulal.ram@aviva.com";
            string MailSub = "ProjectName - Excel file is not available at source location";
            
            string FilePath = Dts.Variables["User::Filepath"].Value.ToString();
            bool FileAvaliablity;

        StartHere:
            {
                if (File.Exists(FilePath))
                    FileAvaliablity = true;           
                else
                {
                    FileAvaliablity = false;
                    SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer);
                    Thread.Sleep(60000);
                    goto StartHere;
                }
            }
            Dts.Variables["User::FileExists"].Value = FileAvaliablity;
			Dts.TaskResult = (int)ScriptResults.Success;
		}


        private void SendMail(string MailFrom, string MailTo, string MailCC, string MailSub, string SMTPServer)
        {
            MailMessage htmlMessage = new MailMessage();
            SmtpClient mySmtpClient = new SmtpClient(SMTPServer);
            StringBuilder MailBody = new StringBuilder();

            MailBody.Append("<html><div style=font-family:Calibri;font-size:11.0pt;color:#000000>");
            MailBody.Append("Hello Team");
            MailBody.Append("</BR></BR>");
            MailBody.Append("Automated system identified that the required file is not avaliable at the location and hence the task have not completed");
            MailBody.Append("</BR></BR>");
            MailBody.Append("<b>NOTE:</b> Automated system will again look for file at the location after 10 minutes and will proceed further with the process if file is available");
            MailBody.Append("</BR></BR>");
            MailBody.Append("Thank you,");
            MailBody.Append("</BR>");
            MailBody.Append("YourReports Team");
            MailBody.Append("</div style></html>");

            htmlMessage = new MailMessage(MailFrom, MailTo, MailSub, MailBody.ToString());
            //htmlMessage.To.Add("babulal.ram@aviva.com");
            //htmlMessage.CC.Add(MailCC.ToString());
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