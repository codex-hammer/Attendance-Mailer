using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Data;
using System.Net.Mail;
using System.Text;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Logging;
namespace AttendanceMailer
{
    class SendMailService
    {
        // This function write log to LogFile.text when some error occurs.      
        public static void WriteErrorLog(Exception ex)
        {
            StreamWriter sw = null;
            try
            {
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + ex.Source.ToString().Trim() + "; " + ex.Message.ToString().Trim());
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }
        // This function write Message to log file.    
        public static void WriteErrorLog(string Message)
        {
            StreamWriter sw = null;
            try
            {
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + Message);
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }

        public void SendmailtoEmployee()
        {
            //var x = 888;
            getEmpAttendance();

        }

        public DataSet getEmplDetails()
        {
            Database _database;
            DbCommand dbCommand = null;
            DataSet Ds = new DataSet();
            try
            {
                Common.CreateLoggingEntry("in Sql Call -", "");
                String spName = ConfigurationManager.AppSettings["MailingStoreProcedure"];

                _database = Common.GetDatabase();
                dbCommand = _database.GetSqlStringCommand("GetEmployeeDetailsList");
                dbCommand.CommandType = CommandType.StoredProcedure;
                dbCommand.CommandTimeout = 0;
                Ds = _database.ExecuteDataSet(dbCommand);
                if (Ds.Tables.Count == 0)
                { Ds = null; }
                else if (Ds.Tables[0].Rows.Count == 0)
                { Ds = null; }

                Common.CreateLoggingEntry("after Sql Call -", "");
            }
            catch (Exception ex)
            {
                Common.CreateLoggingEntry("in Sql Call exception-", "");
                throw ex;
            }
            finally
            {
                dbCommand = null;
                _database = null;
            }
            return Ds;
        }


        public DataSet getAttendance(string empGPN)
        {
            Database _database;
            int month = System.DateTime.Now.Month;
            DbCommand dbCommand = null;
            DataSet Ds = new DataSet();
            try
            {
                Common.CreateLoggingEntry("in Sql Call -", "");
                String spName = ConfigurationManager.AppSettings["MailingStoreProcedure"];

                _database = Common.GetDatabase();
                dbCommand = _database.GetSqlStringCommand("GetAttendanceRecordForGPN");
                _database.AddInParameter(dbCommand, "@empGPN", DbType.String, empGPN);
                _database.AddInParameter(dbCommand, "@month", DbType.String, month);
                dbCommand.CommandType = CommandType.StoredProcedure;
                dbCommand.CommandTimeout = 0;
                Ds = _database.ExecuteDataSet(dbCommand);
                if (Ds.Tables.Count == 0)
                { Ds = null; }
                else if (Ds.Tables[0].Rows.Count == 0)
                { Ds = null; }

                Common.CreateLoggingEntry("after Sql Call -", "");
            }
            catch (Exception ex)
            {
                Common.CreateLoggingEntry("in Sql Call exception-", "");
                throw ex;
            }
            finally
            {
                dbCommand = null;
                _database = null;
            }
            return Ds;
        }


        public void getEmpAttendance()
        {
            DataSet ds = getEmplDetails();
            foreach (DataTable table in ds.Tables)
            {
                foreach (DataRow dr in table.Rows)
                {
                    string empgpn = dr["EmpGPN"].ToString();
                    string empemail = dr["EmpEmail"].ToString();
                    string manageremail = dr["EMPMgrEmail"].ToString();
                    string empname = dr["EmpName"].ToString();
                    DataSet dss = getAttendance(empgpn);
                    if (dss != null)
                    {
                        int count1 = Common.CountCheckIns(dss.Tables[0]);
                        int count2 = Common.CountCheckOuts(dss.Tables[0]);
                        int count3 = Common.CountLowWorkingHours(dss.Tables[0]);
                        dss.Tables[0].Columns.Remove("Whours");
                        Common.ExportDataTableToExcel(dss.Tables[0], empname);
   
                        SendEmail(empemail, manageremail, empname, empgpn, count1, count2, count3);

                        string filepath = ConfigurationManager.AppSettings["filepath"].ToString();
                        Common.DeleteFile(filepath + empname + " (Attendance).xls");

                    }
                }
            }
        }

        // This function contains the logic to send mail.    
        public static void SendEmail(String ToEmail, String mgrEmail, String empname, string empgpn, int count1, int count2, int count3)
        {
            try
            {

                System.Net.Mail.SmtpClient smtpServer = new System.Net.Mail.SmtpClient();
                smtpServer.Host = ConfigurationManager.AppSettings["SMTPServer"].ToString();
                smtpServer.UseDefaultCredentials = false;

                smtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["SMTPUserName"].ToString(), ConfigurationManager.AppSettings["SMTPPassword"].ToString(), ConfigurationManager.AppSettings["SMTPDomain"].ToString());
                smtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpServer.Port = Convert.ToInt32(ConfigurationManager.AppSettings["SMTPPort"]);
                string MailDisplayName = ConfigurationManager.AppSettings["MailDisplayName"].ToString();
                string MailFrom = ConfigurationManager.AppSettings["MailFrom"].ToString();
                string filepath = ConfigurationManager.AppSettings["filepath"].ToString();

                MailMessage MailMsg = new MailMessage();

                System.Net.Mime.ContentType HTMLType = new System.Net.Mime.ContentType("text/html");

                string strBody = Common.createbody(empname, count1, count2, count3);

                MailMsg.BodyEncoding = System.Text.Encoding.Default;


                MailMsg.From = new MailAddress("IN_ITAPP_SVC@IN.EY.COM", MailDisplayName);
                MailMsg.Sender = new MailAddress(MailFrom);
                MailMsg.CC.Add(mgrEmail);
                MailMsg.To.Add(ToEmail);
                MailMsg.Priority = System.Net.Mail.MailPriority.High;
                MailMsg.Subject = "CheckIn/CheckOut Report For month " + (DateTime.Now.Month).ToString();
                MailMsg.Body = strBody;
                MailMsg.Attachments.Add(new Attachment(filepath + empname + " (Attendance).xls"));

                MailMsg.IsBodyHtml = true;
                System.Net.Mail.AlternateView HTMLView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(strBody, HTMLType);

                smtpServer.Send(MailMsg);
                WriteErrorLog("Mail sent successfully!");
                MailMsg.Dispose();
            }
            catch (Exception ex)
            {
                WriteErrorLog(ex.InnerException.Message);
                throw;
            }
        }

    }
}