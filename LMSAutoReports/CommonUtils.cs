using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Data;

namespace LMSAutoReports
{
    public class CommonUtils
    {
        public enum eReportType
        {
            CustomReport = 0, ComplianceReport = 1, ByLearningPathReport = 2, LearningPathDetailsReport = 3, DetailedQuizReport = 4
        }

        public static eReportType ReportType { get; set; }

        public class Frequency
        {
            public string type { get; set; }
            public int day_of_month { get; set; }
            public List<string> day_of_week { get; set; }
        }

        public class EmailReportArgs
        {
            public string report_email_subject { get; set; }
            public string report_name { get; set; }
            public string report_email_message { get; set; }
        }

        // Check if report is set to run weekly/daily.
        public static bool checkReportFrequency(Frequency frequency)
        {
            // If weekly, then check which day of the week is specified for the report.
            if (frequency.type == "weekly")
            {
                // If report is set to run weekly, then check if the day of week specified in the argument list matches the current date.
                string currentDay = DateTime.Now.DayOfWeek.ToString().ToLower();
                if (frequency.day_of_week.Contains(currentDay))
                {
                    return true;
                }
            }
            else if (frequency.type == "monthly")
            {
                // if report is set to run monthly, then check if the day of the month in the argument list matches the current day.
                int currentDayOfMonth = DateTime.Now.Day;
                if (frequency.day_of_month == currentDayOfMonth)
                {
                    return true;
                }
            }
            // If daily then generate the report regardless of the current day.
            else if (frequency.type == "daily")
            {
                return true;
            }
            return false;
        }

        public static int GetLMSID(int OrgID)
        {
            string sql = "SELECT LMSID from Organization "
                   + "WHERE OrganizationID = @OrganizationID";

            var pars = new Dictionary<string, object>();
            pars.Add("@OrganizationID", OrgID);

            DataView dv = Utility.GetDataFromQueryPortal(sql, CommandType.Text, pars);
            DataTable dt = dv.Table;
            if (dt.Rows.Count > 0)
            {
                return Convert.ToInt32(dt.Rows[0]["LMSID"]);
            }
            return 0;
        }

        /// <summary>
        /// Check if the organization has a requirement to have a custom date format other than the defult.
        ///        To Test this use this query, where the OrganizationID is whatever value you have in your local database.
        ///        INSERT INTO [dbo].[OrganizationSettings] ([OrganizationID],[Category],[Setting],[Value]) 
        ///        VALUES (<OrganizationID, int,>,'CS','DF','yyyy-MM-dd');
        /// </summary>
        /// <param name="OrgID"></param>
        /// <returns></returns>
        public static string GetDateFormatFrmOrgSettings(int OrgID)
        {
            string sql = "SELECT Value from OrganizationSettings "
                    + "WHERE OrganizationID = @OrganizationID AND Setting = @Setting AND Category = @Category";

            var pars = new Dictionary<string, object>();
            pars.Add("@OrganizationID", OrgID);
            pars.Add("@Setting", "DF");
            pars.Add("@Category", "CS");

            DataView dv = Utility.GetDataFromQueryPortal(sql, CommandType.Text, pars);
            DataTable dt = dv.Table;
            if (dt.Rows.Count > 0)
            {
                return dt.Rows[0]["Value"].ToString();
            }
            return "MM/dd/yyyy";
        }

        public static bool sendEmail(string strSubject, List<string> toList, string strMessage, string reportName, int orgID, string emailReportLocation = null, bool reportAsAttachment = false)
        {
            string server = System.Configuration.ConfigurationManager.AppSettings["mailServer"];
            string fromAddress = System.Configuration.ConfigurationManager.AppSettings["fromEmail"];

            string account = System.Configuration.ConfigurationManager.AppSettings["mailAccount"];
            string accountpwd = System.Configuration.ConfigurationManager.AppSettings["mailPwd"];

            string reportDownloadLink = System.Configuration.ConfigurationManager.AppSettings["reportDownloadLink"];
            try
            {
                foreach (string to in toList)
                {
                    MailMessage message = new MailMessage(fromAddress, to, strSubject, strMessage);
                    if (reportAsAttachment)
                    {
                        message.Body = strMessage;
                        message.Attachments.Add(new Attachment(emailReportLocation));
                    }
                    else
                    {
                        message.Body = "<a href='" + reportDownloadLink + reportName + "_" + +orgID + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".csv'> Click Here to Download " + "</a><br><br>" + strMessage;                       
                    }

                    message.IsBodyHtml = true;
                    SmtpClient smtpClient = new SmtpClient(server);

                    smtpClient.Port = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["SMTPPort"]);
                    if (accountpwd != string.Empty)
                        smtpClient.Credentials = new System.Net.NetworkCredential(account, accountpwd);
                    smtpClient.Send(message);
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
