using System;
using System.Data;
using System.Collections.Generic;
using System.IO;

namespace LMSAutoReports
{
    public class CourseQuizReportModel
    {
        public class CourseQuizReportArgs
        {
            public List<string> headers { get; set; }
            public List<string> recipients { get; set; }
            public int courseID { get; set; }
            public bool report_as_attachment { get; set; }
            public CommonUtils.Frequency frequency { get; set; }
            public CommonUtils.EmailReportArgs email_report_args { get; set; }            
        }
        public static CourseQuizReportArgs DeserializeQuizReportArgs(string argListString)
        {
            CourseQuizReportArgs args = Newtonsoft.Json.JsonConvert.DeserializeObject<CourseQuizReportArgs>(argListString);
            return args;
        }
    }
    public class CourseQuizReport
    {
        #region Variables

        public int QuestionNo { get; set; }
        public string QuestionName { get; set; }
        public string Correct { get; set; }
        public string Incorrect { get; set; }
        public string Partial { get; set; }
        #endregion

        #region Constructors
        public CourseQuizReport(DataRow dr)
        {
            SetValues(dr);
        }
        #endregion

        private void SetValues(DataRow dr)
        {
            QuestionNo = (int)dr["InteractionID"];
            QuestionName = dr["name"].ToString();
            Correct = dr["correct"].ToString();
            Incorrect = dr["incorrect"].ToString();
            Partial = dr["partial"].ToString();
        }

        private static List<CourseQuizReport> GetQuizReport(int courseID)
        {
            // Generate quiz report using the provided arguments.
            DataView dv = getQuizReport(courseID);

            List<CourseQuizReport> list = new List<CourseQuizReport>();
            foreach (DataRow dr in dv.Table.Rows)
            {
                CourseQuizReport singleReportRow = new CourseQuizReport(dr);
                list.Add(singleReportRow);
            }
            return list;
        }

        // Function to call the stored procedure to generate the quiz report.
        private static DataView getQuizReport(int courseID)
        {
            string sql = "spGetInteractionsReport";
            var pars = new Dictionary<string, object>();

            pars.Add("@CourseOfferingID", courseID);
            return Utility.GetDataFromQueryGadget(sql, CommandType.StoredProcedure, pars);
        }

        public static void createQuizReport(AutoReporting report, string reportFileLocation)
        {
            // Desiralize the list of argments required to run the report. 
            CourseQuizReportModel.CourseQuizReportArgs reportArgs = CourseQuizReportModel.DeserializeQuizReportArgs(report.ReportArgList);

            // Start the report generation process only if the report is set to run on the current day.
            if (!CommonUtils.checkReportFrequency(reportArgs.frequency))
            {
                return;
            }
            else
            {
                int LMSID = CommonUtils.GetLMSID(report.OrganizationID);
                // Get the report using the report arguments specified in the argument list in the auto report table.
                List<CourseQuizReport> gsaReport = CourseQuizReport.GetQuizReport(reportArgs.courseID);
                if (gsaReport.Count > 0)
                {
                    string reportPath = reportFileLocation + reportArgs.email_report_args.report_name.Trim() + "_" + report.OrganizationID + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".csv";
                    writeQuizReportToCSV(gsaReport, reportArgs.headers, reportPath);
                    if (File.Exists(reportPath))
                    {
                        CommonUtils.sendEmail(reportArgs.email_report_args.report_email_subject, reportArgs.recipients, reportArgs.email_report_args.report_email_message,
                                              reportArgs.email_report_args.report_name, report.OrganizationID, reportPath, reportArgs.report_as_attachment);
                    }
                }
            }
        }

        private static void writeQuizReportToCSV(List<CourseQuizReport> reportRows, List<string> reportArgHeaders, string reportLocation)
        {
            List<string> headers = reportArgHeaders;
            using (StreamWriter writer = new StreamWriter(reportLocation))
            {
                // Write the headers to the csv report.
                writer.WriteLine(string.Join(",", headers));
                foreach (CourseQuizReport quizReportRow in reportRows)
                {
                    string[] selectedColumns = new string[headers.Count];
                    for (int i = 0; i < headers.Count; i++)
                    {
                        // Based on the header traverse through the QuizReport row and add to csv.
                        string header = headers[i];
                        switch (header)
                        {
                            case "QuestionNo":
                                selectedColumns[i] = quizReportRow.QuestionNo.ToString();
                                break;
                            case "QuestionName":
                                // Making sure that if a question contains a comma it does not mess-up the report formatting.
                                selectedColumns[i] = $"\"{quizReportRow.QuestionName}\"";
                                break;
                            case "Correct%":
                                selectedColumns[i] = quizReportRow.Correct;
                                break;
                            case "Incorrect%":
                                selectedColumns[i] = quizReportRow.Incorrect;
                                break;
                            case "PartialCorrect%":
                                selectedColumns[i] = quizReportRow.Partial;
                                break;
                        }
                    }
                    writer.WriteLine(string.Join(",", selectedColumns));
                }
            }
        }
    }
}
