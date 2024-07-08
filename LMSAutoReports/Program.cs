using System;
using System.Collections.Generic;

namespace LMSAutoReports
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Getting all reports from the AutoReporting table. Continue further only if any report exists.
            string reportLocation = System.Configuration.ConfigurationManager.AppSettings["reportLocation"];
            string errorLog = System.Configuration.ConfigurationManager.AppSettings["errorLog"];
            List<AutoReporting> reports = AutoReporting.GetAllReports();
            if (reports.Count > 0)
            {
                foreach (AutoReporting report in reports)
                {
                    try
                    {
                        // Function to generate the report.
                        generateReport(report, reportLocation);
                    }
                    catch (Exception ex)
                    {
                        LogMessageToFile(ex.Message, errorLog);
                    }
                }
            }
        }

        public static void generateReport(AutoReporting report, string reportLocation)
        {
            // Swtich statement used to run the report specified in the auto reporting table.
            switch (report.ReportType)
            {
                case CommonUtils.eReportType.CustomReport:
                    // Function to run the Custom Report.
                    CustomReport.createCustomReport(report, reportLocation);                   
                    break;
                case CommonUtils.eReportType.ComplianceReport:
                    break;
                case CommonUtils.eReportType.ByLearningPathReport:
                    // Function to run the Learning Path Report.
                    LearningPathReport.createLPReport(report, reportLocation);
                    break;
                case CommonUtils.eReportType.LearningPathDetailsReport:
                    break;
                case CommonUtils.eReportType.DetailedQuizReport:
                    // Function to run the Detailed Quiz Report.
                    CourseQuizReport.createQuizReport(report, reportLocation);
                    break;
                default:
                    break;
            }
        }

        public static void LogMessageToFile(string msg, string path)
        {
            using (System.IO.StreamWriter sw = System.IO.File.AppendText(path))
            {
                string logLine = System.String.Format("{0:G}: {1}.", System.DateTime.Now, msg);
                sw.WriteLine(logLine);
            }
        }

    }
}
