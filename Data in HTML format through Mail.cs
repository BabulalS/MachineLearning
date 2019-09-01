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
            StringBuilder sbsubject; // = new StringBuilder();
            string strserver = "SMTP.VIA.NOVONET";
            string constr = @"Data Source=UKNWSVUAB268;Initial Catalog=MI_SS_WEB;Integrated Security=SSPI;";
            SqlConnection con = new SqlConnection(constr);
            SqlCommand cmd = new SqlCommand("select * from [vw_ATOM_Tracker_Data] order by owner", con);
            SqlDataAdapter da = new SqlDataAdapter();
            con.Open();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds);
            string mgrid = "";
            string strsendto = "";
            string strfrom = "operationexcellence@aviva.com";
            string strsub = "";
            if (ds.Tables.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                StringBuilder sbh = new StringBuilder();
                sbh.Append("<TR  bgcolor='0070C0'>");
                for (int k = 0; k < ds.Tables[0].Columns.Count - 1; k++)
                {
                    sbh.Append("<TD><font color='#FFFFFF'>");
                    sbh.Append(ds.Tables[0].Columns[k].ToString());
                    sbh.Append("</font></TD>");
                }
                sbh.Append("</TR>");

                if (ds.Tables[0].Rows.Count > 0)
                {
                    strsendto = "";
                    strfrom = "operationexcellence@aviva.com";
                    strsub = "";
                    mgrid = ds.Tables[0].Rows[0]["Email Address"].ToString();
                    sb = new StringBuilder();
                    strsendto = mgrid;
                    strsub = "ATOM - Reports";
                    sb.Append("Hi ");
                    sb.Append(ds.Tables[0].Rows[0]["owner"].ToString());
                    sb.Append(" ,");
                    sb.Append("</BR></BR>");
                    sb.Append("Please find the below list of Report(s) which is/are Due today or Overdue in ATOM  as of today.");
                    sb.Append("</BR></BR>");
                    //sb.Append("Kindly have these updated with immediate effect.");
                    //sb.Append("</BR></BR>");
                    sb.Append("<TABLE border='1' width='1300px'>");
                    sb.Append(sbh.ToString());
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        if (mgrid != ds.Tables[0].Rows[i]["Email Address"].ToString())
                        {
                            sb.Append("</TABLE>");
                            SendMailMessage(strsendto, strfrom, strsub, sb.ToString(), true, strserver);
                            sb = new StringBuilder();
                            //send message code
                            mgrid = ds.Tables[0].Rows[i]["Email Address"].ToString();
                            strsendto = mgrid;
                            strsub = "ATOM - Reports";
                            sb.Append("Hi ");
                            sb.Append(ds.Tables[0].Rows[i]["owner"].ToString());
                            sb.Append(" ,");
                            sb.Append("</BR></BR>");
                            sb.Append("Please find the below list of Report(s) which is/are Due today or Overdue in ATOM  as of today.");                            
                            sb.Append("</BR></BR>");
                            sb.Append("<TABLE border='1' width='1300px'>");
                            sb.Append(sbh.ToString());
                            if (Convert.ToDateTime(ds.Tables[0].Rows[i]["Required By"]) <= DateTime.Today.AddDays(-1))
                            {
                                sb.Append("<TR bgcolor='#FF0000'>");  // red color

                                sb.Append("<TD width='10%'> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Reference"].ToString());
                                sb.Append("</font> </TD>");
                                sb.Append("<TD  width='20%'> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["job name"].ToString());
                                sb.Append("</font> </TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Status"].ToString());
                                sb.Append("</font> </TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Frequency"].ToString());
                                sb.Append("</font> </TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Required By"].ToString());
                                sb.Append("</font> </TD>");                                
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Owner"].ToString());
                                sb.Append("</font> </TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Deputy"].ToString());
                                sb.Append("</font> </TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Reporting Team"].ToString());
                                sb.Append("</font> </TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Time to Deliver"].ToString());
                                sb.Append("</font> </TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Priority"].ToString());
                                sb.Append("</font> </TD>");
                                //sb.Append("<TD> <font color='#FFFFFF'>");
                                //sb.Append(ds.Tables[0].Rows[i]["JC Link"].ToString());
                                //sb.Append(" </font> </TD>");
                                //sb.Append("<TD> <font color='#FFFFFF'>");
                                //sb.Append(ds.Tables[0].Rows[i]["Result Link"].ToString());
                                //sb.Append("</font> </TD>");
                                //sb.Append("<TD> <font color='#FFFFFF'>");
                                //sb.Append(ds.Tables[0].Rows[i]["AW Link"].ToString());
                                //sb.Append("</font> </TD>"); 
                            }
                            else
                            {
                                sb.Append("<TR  bgcolor='FFDF79'> ");  // yellow color

                                sb.Append("<TD width='10%'>");
                                sb.Append(ds.Tables[0].Rows[i]["Reference"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD  width='20%'>");
                                sb.Append(ds.Tables[0].Rows[i]["job name"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Status"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Frequency"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Required By"].ToString());
                                sb.Append("</TD>");                                
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Owner"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Deputy"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Reporting Team"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Time to Deliver"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Priority"].ToString());
                                sb.Append("</TD>");
                                //sb.Append("<TD>");
                                //sb.Append(ds.Tables[0].Rows[i]["JC Link"].ToString());                                
                                //sb.Append("</TD>");
                                //sb.Append("<TD>");
                                //sb.Append(ds.Tables[0].Rows[i]["Result Link"].ToString());
                                //sb.Append("</TD>");
                                //sb.Append("<TD>");
                                //sb.Append(ds.Tables[0].Rows[i]["AW Link"].ToString());
                                //sb.Append("</TD>"); 
                            }  
                                                     

                            sb.Append("</TR>");
                        }
                        else   // main if else
                        {
                            // sb.Append("<TR width='25%'>");
                            if (Convert.ToDateTime(ds.Tables[0].Rows[i]["Required By"]) <= DateTime.Today.AddDays(-1))
                            {
                                sb.Append("<TR  bgcolor='#FF0000'>");  /// red color
                                sb.Append("<TD width='10%'> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Reference"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD  width='20%'> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["job name"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Status"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Frequency"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Required By"].ToString());
                                sb.Append("</font></TD>");
                                
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Owner"].ToString());
                                sb.Append("</font> </TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Deputy"].ToString());
                                sb.Append("</font> </TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Reporting Team"].ToString());
                                sb.Append("</font> </TD>");
                                sb.Append("<TD><font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Time to Deliver"].ToString());
                                sb.Append("</font> </TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Priority"].ToString());
                                sb.Append("</font> </TD>");
                                //sb.Append("<TD> <font color='#FFFFFF'>");
                                //sb.Append(ds.Tables[0].Rows[i]["JC Link"].ToString());
                                //sb.Append("</font> </TD>");
                                //sb.Append("<TD> <font color='#FFFFFF'>");
                                //sb.Append(ds.Tables[0].Rows[i]["Result Link"].ToString());
                                //sb.Append("</font> </TD>");
                                //sb.Append("<TD> <font color='#FFFFFF'>");
                                //sb.Append(ds.Tables[0].Rows[i]["AW Link"].ToString());
                                //sb.Append("</font> </TD>"); 
                            }
                            else
                            {
                                sb.Append("<TR  bgcolor='FFDF79'> ");  // yellow color
                                sb.Append("<TD width='10%'>");
                                sb.Append(ds.Tables[0].Rows[i]["Reference"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD  width='20%'>");
                                sb.Append(ds.Tables[0].Rows[i]["job name"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Status"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Frequency"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Required By"].ToString());
                                sb.Append("</TD>");                                
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Owner"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Deputy"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Reporting Team"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Time to Deliver"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Priority"].ToString());
                                sb.Append("</TD>");
                                //sb.Append("<TD>");
                                //sb.Append(ds.Tables[0].Rows[i]["JC Link"].ToString());                                
                                //sb.Append("</TD>");
                                //sb.Append("<TD>");
                                //sb.Append(ds.Tables[0].Rows[i]["Result Link"].ToString());
                                //sb.Append("</TD>");
                                //sb.Append("<TD>");
                                //sb.Append(ds.Tables[0].Rows[i]["AW Link"].ToString());
                                //sb.Append("</TD>"); 
                            }
                                                      
                            sb.Append("</TR>");
                        }
                        if (i == ds.Tables[0].Rows.Count - 1)
                        {
                            sb.Append("</TABLE>");
                            SendMailMessage(strsendto, strfrom, strsub, sb.ToString(), true, strserver);
                        }
                    }
                }

            //    //test
               //   SendMailToOE(sbh, ds);
            //    //end test
            }
            con.Close();

           // test
            SendMailToManagerList();

             //end test

            Dts.TaskResult = (int)ScriptResults.Success;
        }

        private void SendMailToOE(StringBuilder sbh, DataSet ds)
        {
            StringBuilder sbOE = new StringBuilder();
            sbOE.Append("<TABLE border='1' width='1300px'>");
            sbOE.Append(sbh.ToString());

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                sbOE.Append("<TR >");
                sbOE.Append("<TD width='10%'>");
                sbOE.Append(ds.Tables[0].Rows[i]["Reference"].ToString());
                sbOE.Append("</TD>");
                sbOE.Append("<TD  width='20%'>");
                sbOE.Append(ds.Tables[0].Rows[i]["job name"].ToString());
                sbOE.Append("</TD>");
                sbOE.Append("<TD>");
                sbOE.Append(ds.Tables[0].Rows[i]["Status"].ToString());
                sbOE.Append("</TD>");
                sbOE.Append("<TD>");
                sbOE.Append(ds.Tables[0].Rows[i]["Frequency"].ToString());
                sbOE.Append("</TD>");
                sbOE.Append("<TD>");
                sbOE.Append(ds.Tables[0].Rows[i]["Required By"].ToString());
                sbOE.Append("</TD>");
                sbOE.Append("<TD>");
                sbOE.Append(ds.Tables[0].Rows[i]["Time Deliver"].ToString());
                sbOE.Append("</TD>");
                sbOE.Append("<TD>");
                sbOE.Append(ds.Tables[0].Rows[i]["Owner"].ToString());
                sbOE.Append("</TD>");
                sbOE.Append("<TD>");
                sbOE.Append(ds.Tables[0].Rows[i]["Deputy"].ToString());
                sbOE.Append("</TD>");
                sbOE.Append("<TD>");
                sbOE.Append(ds.Tables[0].Rows[i]["Reporting Team"].ToString());
                sbOE.Append("</TD>");
                sbOE.Append("<TD>");
                sbOE.Append(ds.Tables[0].Rows[i]["Priority"].ToString());
                sbOE.Append("</TD>");
                sbOE.Append("<TD>");
                sbOE.Append(ds.Tables[0].Rows[i]["JC Link"].ToString());
                sbOE.Append("</TD>");
                sbOE.Append("<TD>");
                sbOE.Append(ds.Tables[0].Rows[i]["Result Link"].ToString());
                sbOE.Append("</TD>");
                sbOE.Append("<TD>");
                sbOE.Append(ds.Tables[0].Rows[i]["AW Link"].ToString());
                sbOE.Append("</TD>");
                //sb.Append("<TD>");
                //sb.Append(ds.Tables[0].Rows[i]["Distribution_List"].ToString());
                //sb.Append("</TD>");

                sbOE.Append("</TR>");
            }

            sbOE.Append("</TABLE'>");

            SendMailMessage(@"operationexcellence@mgd.aviva.com", @"yrrpts1@aviva.com", "ATOM - Reports", sbOE.ToString(), true, "SMTP.VIA.NOVONET");


        }

        private void SendMailMessage(string SendTo, string From, string Subject, string Body, bool IsBodyHtml, string Server)
        {
            try
            {
                MailMessage htmlMessage = new MailMessage();

                Body = Body + @"</Br></BR>Note: This is Autogenerated mail ... Please do not reply";

                Body = Body + @"</Br></BR> Color Code : </BR> Yellow : Current Reports  </BR> Red : Due Reports ";

                SmtpClient mySmtpClient;
                //  SendTo = @"PRAJAPV@aviva.com";
                htmlMessage = new MailMessage(From, SendTo, Subject, Body);
                //htmlMessage.CC.Add(cc.ToString());
                htmlMessage.IsBodyHtml = IsBodyHtml;
                mySmtpClient = new SmtpClient(Server);
                mySmtpClient.Credentials = CredentialCache.DefaultNetworkCredentials;
                mySmtpClient.Send(htmlMessage);
            }
            catch
            {
                //MessageBox.Show(SendTo);
            }
        }

        private void SendMailToManagerList()
        {
            StringBuilder sbsubject; // = new StringBuilder();
            string strserver = "SMTP.VIA.NOVONET";
            string constr = @"Data Source=UKNWSVUAB268;Initial Catalog=MI_SS_WEB;Integrated Security=SSPI;";
            SqlConnection con = new SqlConnection(constr);
            SqlCommand cmd = new SqlCommand("select * from  dbo.VW_ATOM_Manager_Mail_List order by [manager name] , [required by] ,[Reporting Team]", con);
            SqlDataAdapter da = new SqlDataAdapter();
            con.Open();
            da.SelectCommand = cmd;
            DataSet ds = new DataSet();
            da.Fill(ds);
            string mgrid = "";
            string strsendto = "";
            string strfrom = "operationexcellence@aviva.com";
            string strsub = "";
            if (ds.Tables.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                StringBuilder sbh = new StringBuilder();
                sbh.Append("<TR  bgcolor='0070C0'>");  // bg color blue
                for (int k = 2; k < ds.Tables[0].Columns.Count ; k++)
                {
                    sbh.Append("<TD><font color='#FFFFFF'>");  // font white colore
                    sbh.Append(ds.Tables[0].Columns[k].ToString());
                    sbh.Append("</font></TD>");
                }
                sbh.Append("</TR>");

                if (ds.Tables[0].Rows.Count > 0)
                {
                    strsendto = "";
                    strfrom = "operationexcellence@aviva.com";
                    strsub = "";
                    mgrid = ds.Tables[0].Rows[0]["Email Address"].ToString();
                    sb = new StringBuilder();
                    strsendto = mgrid;
                    strsub = "ATOM - Reports";
                    sb.Append("Hi ");
                    sb.Append(ds.Tables[0].Rows[0]["manager name"].ToString());
                    sb.Append(" ,");
                    sb.Append("</BR></BR>");
                    sb.Append("Please find the below list of Report(s) which is/are Due today or Overdue in ATOM  as of today.");
                    sb.Append("</BR></BR>");
                    //sb.Append("Kindly have these updated with immediate effect.");
                    //sb.Append("</BR></BR>");
                    sb.Append("<TABLE border='1' width='1300px'>");
                    sb.Append(sbh.ToString());
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        if (mgrid != ds.Tables[0].Rows[i]["Email Address"].ToString())
                        {
                            sb.Append("</TABLE>");
                            SendMailMessage(strsendto, strfrom, strsub, sb.ToString(), true, strserver);
                            sb = new StringBuilder();
                            //send message code
                            mgrid = ds.Tables[0].Rows[i]["Email Address"].ToString();
                            strsendto = mgrid;
                            strsub = "ATOM - Reports";
                            sb.Append("Hi ");
                            sb.Append(ds.Tables[0].Rows[i]["manager name"].ToString());
                            sb.Append(" ,");
                            sb.Append("</BR></BR>");
                            sb.Append("Please find the below list of Report(s) which is/are Due today or Overdue in ATOM  as of today.");
                            //sb.Append("</BR></BR>");
                            //sb.Append("Kindly have these updated with immediate effect.");
                            sb.Append("</BR></BR>");
                            sb.Append("<TABLE border='1' width='1300px'>");
                            sb.Append(sbh.ToString());
                            if (Convert.ToDateTime(ds.Tables[0].Rows[i]["Required By"]) <= DateTime.Today.AddDays(-1))
                            {
                                sb.Append("<TR bgcolor='#FF0000'>");  // red color

                                // for font white
                                sb.Append("<TD width='10%'> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Reference"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD  width='20%'> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["job name"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["owner"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["deputy"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Status"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Frequency"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Required By"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Reporting Team"].ToString());
                                sb.Append("</font></TD>");

                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Time to Deliver"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["priority"].ToString());
                                sb.Append("</font></TD>");  
                            }
                            else
                            {
                                sb.Append("<TR  bgcolor='FFDF79'> ");  // yellow color

                                sb.Append("<TD width='10%'>");
                                sb.Append(ds.Tables[0].Rows[i]["Reference"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD  width='20%'>");
                                sb.Append(ds.Tables[0].Rows[i]["job name"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["owner"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["deputy"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Status"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Frequency"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Required By"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Reporting Team"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Time to Deliver"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["priority"].ToString());
                                sb.Append("</TD>"); 

                            }
                                                   
                           
                            sb.Append("</TR>");
                        }
                        else
                        {
                            // sb.Append("<TR width='25%'>");
                            if (Convert.ToDateTime(ds.Tables[0].Rows[i]["Required By"]) <= DateTime.Today.AddDays(-1))
                            {
                                sb.Append("<TR  bgcolor='#FF0000'>"); // red color
                                // for font white
                                sb.Append("<TD width='10%'> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Reference"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD  width='20%'> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["job name"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD> <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["owner"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["deputy"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Status"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Frequency"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Required By"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Reporting Team"].ToString());
                                sb.Append("</font></TD>");

                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["Time to Deliver"].ToString());
                                sb.Append("</font></TD>");
                                sb.Append("<TD>  <font color='#FFFFFF'>");
                                sb.Append(ds.Tables[0].Rows[i]["priority"].ToString());
                                sb.Append("</font></TD>");
                            }
                            else
                            {
                                sb.Append("<TR  bgcolor='FFDF79'> ");  // yellow color
                                sb.Append("<TD width='10%'>");
                                sb.Append(ds.Tables[0].Rows[i]["Reference"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD  width='20%'>");
                                sb.Append(ds.Tables[0].Rows[i]["job name"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["owner"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["deputy"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Status"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Frequency"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Required By"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Reporting Team"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["Time to Deliver"].ToString());
                                sb.Append("</TD>");
                                sb.Append("<TD>");
                                sb.Append(ds.Tables[0].Rows[i]["priority"].ToString());
                                sb.Append("</TD>");
                            }
                                                      
                            sb.Append("</TR>");
                        }
                        if (i == ds.Tables[0].Rows.Count - 1)
                        {
                            sb.Append("</TABLE>");
                            SendMailMessage(strsendto, strfrom, strsub, sb.ToString(), true, strserver);
                        }
                    }
                }

            }
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