using System;
using System.Data;
using System.Collections.Generic;
using System.IO;

namespace LMSAutoReports
{
    public class LPReportModel
    {
        public class LPReportArgs
        {
            public List<string> headers { get; set; }
            public List<string> curriculum_ids { get; set; }
            public int group_id { get; set; }
            public bool include_inactive { get; set; }
            public bool report_as_attachment { get; set; }
            public List<string> recipients { get; set; }
            public CommonUtils.Frequency frequency { get; set; }
            public CommonUtils.EmailReportArgs email_report_args { get; set; }          
        }

        public static LPReportArgs DeserializeLPReportArgs(string argListString)
        {
            LPReportArgs args = Newtonsoft.Json.JsonConvert.DeserializeObject<LPReportArgs>(argListString);
            return args;
        }
    }
    public class LearningPathReport
    {
        #region Variables

        public string UserName;
        public string FirstName;
        public string LastName;
        public string GroupName;
        public string JobName;
        public string CompletionStatus;
        public string SuccessStatus;
        public decimal AverageScore;
        public string ParentGroup;
        public decimal CompletionPercent;

        #endregion

        #region Constructors
        public LearningPathReport(DataRow dr)
        {
            SetValues(dr);
        }
        #endregion
        private void SetValues(DataRow dr)
        {
            UserName = dr["UserName"].ToString();
            FirstName = dr["Name_N_Given"].ToString();
            LastName = dr["Name_N_Family"].ToString();
            GroupName = dr["AreaName"].ToString();
            JobName = dr["JobTitle"] == DBNull.Value ? String.Empty : dr["JobTitle"].ToString();
            CompletionStatus = dr["completion_status"].ToString();
            SuccessStatus = dr["success_status"].ToString();
            ParentGroup = dr["ParentArea"] == DBNull.Value ? String.Empty : dr["ParentArea"].ToString();
            CompletionPercent = dr["CompletionPercent"] == DBNull.Value ? 0 : (decimal)dr["CompletionPercent"];
        }

        private static List<LearningPathReport> GetLearningPathReport(string curriculumids, int groupid, int lmsid, bool includeInactive)
        {
            // Generate learning path report using the provided arguments.
            DataView dv = getLPReportData(curriculumids, groupid, lmsid, includeInactive);
            List<LearningPathReport> list = new List<LearningPathReport>();
            foreach (DataRow dr in dv.Table.Rows)
            {
                LearningPathReport singleReportRow = new LearningPathReport(dr);
                list.Add(singleReportRow);
            }
            return list;
        }

        // Function to call the stored procedure to generate the learning path report.
        private static DataView getLPReportData(string curriculumid, int groupid, int lmsid, bool includeInactive)
        {
            string sql = "spGetCurriculumAttemptsByID";
            var pars = new Dictionary<string, object>();

            pars.Add("@CurriculumID", curriculumid);
            pars.Add("@AreaID", groupid);
            pars.Add("@LMSID", lmsid);
            pars.Add("@IncludeInactive", includeInactive);

            return Utility.GetDataFromQueryPortal(sql, CommandType.StoredProcedure, pars);
        }

        private static string getCurriculumNameByID(int curriculumID)
        {
            string sql = "SELECT Name from Curriculum "
                   + "WHERE CurriculumID = @CurriculumID";

            var pars = new Dictionary<string, object>();
            pars.Add("@CurriculumID", curriculumID);

            DataView dv = Utility.GetDataFromQueryGadget(sql, CommandType.Text, pars);
            DataTable dt = dv.Table;
            if (dt.Rows.Count > 0)
            {
                return dt.Rows[0]["Name"].ToString();
            }
            return string.Empty;
        }

        public static void createLPReport(AutoReporting report, string reportFileLocation)
        {
            // Desiralize the list of argments required to run the report. 
            LPReportModel.LPReportArgs reportArgs = LPReportModel.DeserializeLPReportArgs(report.ReportArgList);

            // Start the report generation process only if the report is set to run on the current day.
            if (!CommonUtils.checkReportFrequency(reportArgs.frequency))
            {
                return;
            }
            else
            {
                int LMSID = CommonUtils.GetLMSID(report.OrganizationID);
                string reportPath = reportFileLocation + reportArgs.email_report_args.report_name.Trim() + "_" + report.OrganizationID + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".csv";
                bool headersWritten = false;

                // Loop to include multiple LPs for a single report.
                foreach (string currID in reportArgs.curriculum_ids)
                {
                    // Get the report using the report arguments specified in the argument list in the auto report table.
                    List<LearningPathReport> lpReport = LearningPathReport.GetLearningPathReport(currID, reportArgs.group_id, LMSID, reportArgs.include_inactive);
                    if (lpReport.Count > 0)
                    {
                        string curriculumName = getCurriculumNameByID(Convert.ToInt32(currID));                      
                        writeLPReportToCSV(lpReport, reportArgs.headers, reportPath, curriculumName, ref headersWritten);           
                    }
                }
                if (File.Exists(reportPath))
                {
                    CommonUtils.sendEmail(reportArgs.email_report_args.report_email_subject, reportArgs.recipients, reportArgs.email_report_args.report_email_message,
                                          reportArgs.email_report_args.report_name, report.OrganizationID, reportPath, reportArgs.report_as_attachment);
                }
            }

        }

        private static void writeLPReportToCSV(List<LearningPathReport> reportRows, List<string> reportArgHeaders, string reportLocation, string curriculumName, ref bool headersWritten)
        {
            List<string> headers = reportArgHeaders;

            // Write headers if they haven't been written yet.
            if (!headersWritten)
            {
                using (StreamWriter writer = new StreamWriter(reportLocation))
                {
                    writer.WriteLine(string.Join(",", headers));
                }
                headersWritten = true;
            }
            using (StreamWriter writer = new StreamWriter(reportLocation, true))
            {
                foreach (LearningPathReport lpReportRow in reportRows)
                {
                    string[] selectedColumns = new string[headers.Count];
                    for (int i = 0; i < headers.Count; i++)
                    {
                        // Based on the header traverse through the lpReport row and add to csv.
                        string header = headers[i];
                        switch (header)
                        {
                            case "CurriculumName":
                                selectedColumns[i] = "\"" + curriculumName.Replace("\"", "\"\"") + "\"";
                                break;
                            case "UserName":
                                selectedColumns[i] = lpReportRow.UserName;
                                break;
                            case "FirstName":
                                selectedColumns[i] = lpReportRow.FirstName;
                                break;
                            case "LastName":
                                selectedColumns[i] = lpReportRow.LastName;
                                break;
                            case "Group":
                                selectedColumns[i] = lpReportRow.GroupName;
                                break;
                            case "Job":
                                selectedColumns[i] = lpReportRow.JobName;
                                break;
                            case "CompletionStatus":
                                selectedColumns[i] = lpReportRow.CompletionStatus;
                                break;
                            case "SuccessStatus":
                                selectedColumns[i] = lpReportRow.SuccessStatus;
                                break;
                            case "ParentGroup":
                                selectedColumns[i] = lpReportRow.ParentGroup;
                                break;
                            case "CompletionPercent":
                                selectedColumns[i] = Math.Round(lpReportRow.CompletionPercent,2).ToString() + "%";
                                break;
                        }
                    }
                    writer.WriteLine(string.Join(",", selectedColumns));
                }
            }
        }
    }
}
