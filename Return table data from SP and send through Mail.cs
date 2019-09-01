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
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Data.SqlClient;
#endregion

namespace ST_95f473e362c44ba9aeb60f9e842f0ebe
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
            string MailSub = "Telephony Logs - Portal and CallManager System";
            string PortalSQLQuery = Dts.Variables["User::PortalSQLQuery"].Value.ToString();
            string CMSQLQuery = Dts.Variables["User::CMProc"].Value.ToString();
            string MailFrom = "yrrpts1@aviva.com";
            string MailTo = "babulal.ram@aviva.com";
            string MailCC = "babulal.ram@aviva.com";
            string CMBody;
            string PortalBody;
            StringBuilder MailBody = new StringBuilder();

            /////////////////   Code for Portal Counts /////////////////////////////////////////   

            PortalBody = "";
            PortalBody += "</html><head></head><body>Hello Team,<br/><br/>";
            PortalBody += "Please find the below table counts from Portal Source<br/><br/>";
            PortalBody += "<table cellpadding='5' cellspacing='0' style='font-size:12px;font-family: Arial;' border='1'>";
            PortalBody += "<tr Align=Center style='background-color: #D3D3D3;align: Center;'><td><b>Date<b></td><td><b>RCD_All Count<b></td><td><b>TCD_All Count<b></td>";
            PortalBody += "<td><b>RCDCount<b></td><td><b>TCDCount<b></td>";

            string PConnectionInfo = @"Data Source=10.201.225.94,1433;Network Library=DBMSSOCN;Initial Catalog=Portal;User Id=khanna4;password=k1eWRL_2-?sw7aTr";
            SqlConnection PServerConn = new SqlConnection(PConnectionInfo);
            SqlCommand PSQLQuery = new SqlCommand(PortalSQLQuery, PServerConn);
            PSQLQuery.CommandType = CommandType.Text;
            PSQLQuery.CommandTimeout = 0;

            try
            {
                PServerConn.Open();

                if (PServerConn.State == ConnectionState.Open)
                {
                    SqlDataReader objDataReader = PSQLQuery.ExecuteReader();
                    while (objDataReader.Read())
                    {
                        PortalBody += "<tr>";
                        PortalBody += "<td Align=Center>" + Convert.ToDateTime(objDataReader["Date"]).ToShortDateString() + "</td>";
                        PortalBody += "<td Align=Center>" + objDataReader["RCD_ALLCount"].ToString() + "</td>";
                        PortalBody += "<td Align=Center>" + objDataReader["TCD_ALLCount"].ToString() + "</td>";
                        PortalBody += "<td Align=Center>" + objDataReader["RCDCount"].ToString() + "</td>";
                        PortalBody += "<td Align=Center>" + objDataReader["TCDCount"].ToString() + "</td>";
                        PortalBody += "</tr>";
                    }
                    PortalBody += "</table></body></html>";
                    PServerConn.Close();
                    
                }
            }

            catch (Exception ex)
            {
                PortalBody += "</table></body></html>";
                PortalBody += "An Error has occured while processing counts from Portal source.</BR>";
                PortalBody +=  "Error Message="+ ex.ToString();
            }

            MailBody.Append(PortalBody.ToString());
            MailBody.Append("</BR></BR>");

         
            /////////////////   Code for Call Manager Counts /////////////////////////////////////////           
          
          
            CMBody = "";
            CMBody += "</html><head></head><CMBody>...and continuation to above below are the counts from CallManager tables<br/><br/>";
            CMBody += "<table cellpadding='5' cellspacing='0' style='font-size:12px;font-family: Arial;' border='1'>";
            CMBody += "<tr Align=Center style='background-color: #D3D3D3;'><td><b>Date<b></td><td><b>WeekDay<b></td><td><b>RouterCallKeyDay<b></td><td><b>RCD<b></td><td><b>TCD<b></td>";
            CMBody += "<td><b>RCD_All<b></td><td><b>TCD_All<b></td><td><b>CallXDetail<b></td><td><b>TCD_OB_New<b></td><td><b>AST<b></td>";
            CMBody += "<td><b>AgentEvent<b></td><td><b>VCA<b></td><td><b>AgentHalfHour<b></td>";

            string ConnectionInfo = @"Data Source=MissDBServer01;Initial Catalog=Callmanager;Integrated Security=SSPI;";
            SqlConnection ServerConn = new SqlConnection(ConnectionInfo);
            SqlCommand SQLQuery = new SqlCommand(CMSQLQuery, ServerConn);
            SQLQuery.CommandType = CommandType.Text;
            SQLQuery.CommandTimeout = 0;

            try
            {
                ServerConn.Open();

                if (ServerConn.State == ConnectionState.Open)
                {
                    SqlDataReader objDataReader = SQLQuery.ExecuteReader();
                    while (objDataReader.Read())
                    {
                        CMBody += "<tr>";
                        CMBody += "<td Align=Center>" + Convert.ToDateTime(objDataReader["Date"]).ToShortDateString() + "</td>";
                        CMBody += "<td Align=Center>" + objDataReader["WeekDay"].ToString() + "</td>";
                        CMBody += "<td Align=Center>" + objDataReader["RouterCallKeyDay"].ToString() + "</td>";
                        CMBody += "<td Align=Center>" + objDataReader["RCD"].ToString() + "</td>";
                        CMBody += "<td Align=Center>" + objDataReader["TCD"].ToString() + "</td>";
                        CMBody += "<td Align=Center>" + objDataReader["RCD_All"].ToString() + "</td>";
                        CMBody += "<td Align=Center>" + objDataReader["TCD_All"].ToString() + "</td>";
                        CMBody += "<td Align=Center>" + objDataReader["CallXDetail"].ToString() + "</td>";
                        CMBody += "<td Align=Center>" + objDataReader["TCD_OB_New"].ToString() + "</td>";
                        CMBody += "<td Align=Center>" + objDataReader["AST"].ToString() + "</td>";
                        CMBody += "<td Align=Center>" + objDataReader["AgentEvent"].ToString() + "</td>";
                        CMBody += "<td Align=Center>" + objDataReader["VCA"].ToString() + "</td>";
                        CMBody += "<td Align=Center>" + objDataReader["AgentHalfHour"].ToString() + "</td>";
                        CMBody += "</tr>";
                    }
                    CMBody += "</table></body></html>";
                    ServerConn.Close();
                   
                }
            }

            catch (Exception ex)
            {
                CMBody += "</table></body></html>";
                CMBody += "An Error has occured while processing Callmanager Logs.</BR>";
                CMBody += "Error Message="+ ex.ToString()+"</BR>";
            }

            MailBody.Append(CMBody.ToString());

            SendMail(MailFrom, MailTo, MailCC, MailSub, SMTPServer, MailBody);
          
         Dts.TaskResult = (int)ScriptResults.Success;
        }

        private void SendMail(string MailFrom, string MailTo, string MailCC, string MailSub, string SMTPServer, StringBuilder MailBody)
        {
            MailMessage htmlMessage = new MailMessage();
            SmtpClient mySmtpClient = new SmtpClient(SMTPServer);
            MailBody.Append("</BR></BR>");
            MailBody.Append("Thank you,</BR>");
            MailBody.Append("YourReports");
            htmlMessage.Body = MailBody.ToString();
            htmlMessage = new MailMessage(MailFrom, MailTo, MailSub, htmlMessage.Body.ToString());
            htmlMessage.IsBodyHtml = true;
            mySmtpClient.Credentials = CredentialCache.DefaultNetworkCredentials;
            mySmtpClient.Send(htmlMessage);
        }      
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