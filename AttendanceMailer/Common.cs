using Microsoft.Practices.EnterpriseLibrary.Common.Configuration;
using Microsoft.Practices.EnterpriseLibrary.Configuration.Design;
using Microsoft.Practices.EnterpriseLibrary.Data;
using Microsoft.Practices.EnterpriseLibrary.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.IO;
using System.Configuration;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttendanceMailer
{
    public class Common
    {
        public static Database GetDatabase()
        {
            IConfigurationSource ConfigurationSource = default(IConfigurationSource);
            ConfigurationSource = new SystemConfigurationSource();
            DatabaseProviderFactory Factory = new DatabaseProviderFactory(ConfigurationSource);
            return Factory.Create("SAFConnection");
        }
        internal static TransactionScope SetTransScope()
        {
            TransactionScope tranScope = null;
            System.Transactions.TransactionOptions tranOpt = default(System.Transactions.TransactionOptions);
            tranOpt.Timeout = new System.TimeSpan(0, 30, 0);
            tranScope = new TransactionScope(TransactionScopeOption.Required, tranOpt);
            return tranScope;
        }

        public static void CreateLoggingEntry(String message1, String message2)
        {
            LogEntry logEntry = new LogEntry();
            if (!String.IsNullOrEmpty(message1.Trim()))
            {
                logEntry.Message = " Message1:" + message1 + " <br> Message2 :" + message2;
                logEntry.TimeStamp = DateTime.Now;
                Logger.Write(logEntry);
            }
        }

        public static void ExportDataTableToExcel(DataTable table, string empname)
        {
            string filepath = ConfigurationManager.AppSettings["filepath"].ToString();
            string filename = filepath + empname + " (Attendance).xls";
            table.WriteXml(filename);
        }

        public static void DeleteFile(string filename)
        {
            if (File.Exists(filename)) { File.Delete(filename); }
        }

        public static int CountCheckIns(DataTable table)
        {
            int count = 0;
            foreach (DataRow row in table.Rows)
            {
                var x = row["CheckInTime"];
                if (x.ToString().Length == 0) count++;

            }
            return count;
        }

        public static int CountCheckOuts(DataTable table)
        {
            int count = 0;
            foreach (DataRow row in table.Rows)
            {
                var x = row["CheckOutTime"];
                if (x.ToString().Length == 0) count++;

            }

            return count;
        }

        public static int CountLowWorkingHours(DataTable table)
        {
            int count = 0;
            foreach (DataRow row in table.Rows)
            {
                string x = row["Whours"].ToString();
                int z;
                bool f = Int32.TryParse(x, out z);
                if (z < 8)
                {
                  count++;
                 }
                }
            return count;
        }

        public static string createbody(string empname, int count1, int count2, int count3)
        {
            string body = "";
            body = body + "Dear " + empname + ",<br> Please find attatched excel sheet for the report of your checkIn and checkOut Timings for this month.<br><br><br> ";
            if (count1 > 0 || count2 > 0 || count3 > 0)
            {
                body = body + "Below are some discrepancies:" +
                              "<table cellpadding='4' border='2'>" +
                                "<tr><th>S.No</th><th>Discrepancy</th><th>Count</th></tr>" +
                                "<tr><td>1</td><td>No. Of Days CheckIn Missed</td><td>" + count1.ToString() + "</td></tr>" +
                                "<tr><td>2</td><td>No. Of Days CheckOut Missed</td><td>" + count2.ToString() + "</td></tr>" +
                                "<tr><td>3</td><td>No. Of Days Working Hours less than 8</td><td>" + count3.ToString() + "</td></tr></table>";
            }
            return body;
        }
    }
}
