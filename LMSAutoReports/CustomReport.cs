using System;
using System.Data;
using System.Collections.Generic;
using System.IO;

namespace LMSAutoReports
{
    public class CustomReportModel
    {
        public class CustomReportArgs
        {
            public List<string> headers { get; set; }
            public int completion_status { get; set; }
            public int success_status { get; set; }
            public int course { get; set; }
            public bool most_recent_attempts { get; set; }
            public bool include_not_attempted { get; set; }
            public List<string> recipients { get; set; }
            public int completion_days_back { get; set; }
            public int hire_days_back { get; set; }
            public CommonUtils.Frequency frequency { get; set; }
            public CommonUtils.EmailReportArgs email_report_args { get; set; }
            public int group { get; set; }
            public string category { get; set; }
            public string meta { get; set; }
            public bool inactive { get; set; }
            public bool report_as_attachment { get; set; }
            public List<int> courses_not_required { get; set; }
            public List<int> groups_not_required { get; set; }
        }       

        public static CustomReportArgs DeserializeCustomReportArgs(string argListString)
        {
            CustomReportArgs args = Newtonsoft.Json.JsonConvert.DeserializeObject<CustomReportArgs>(argListString);
            return args;
        }
    }
    public class CustomReport
    {
        #region Variables

        public string ActivityName;
        public int CourseID;
        public string UserName;
        public string FirstName;
        public string LastName;
        public string Email;
        public string GroupName;
        public int AreaID;
        public string JobName;
        public string EmployeeID;
        public string CompletionStatus;
        public string SuccessStatus;
        public DateTime? CompletionDate;
        public decimal Score;
        public string ParentGroup;
        public string Category;
        public DateTime? HireDate;
        public string Manager;

        #endregion

        #region Constructors
        public CustomReport(DataRow dr)
        {
            SetValues(dr);
        }
        #endregion
        private void SetValues(DataRow dr)
        {
            ActivityName = dr["CourseName"].ToString();
            CourseID = (int)dr["CourseID"];
            UserName = dr["UserName"].ToString();
            FirstName = dr["Name_N_Given"].ToString();
            LastName = dr["Name_N_Family"].ToString();
            Email = dr["Email"] == DBNull.Value ? String.Empty : dr["Email"].ToString();
            GroupName = dr["AreaName"].ToString();
            AreaID = (int)dr["AreaID"];
            JobName = dr["JobTitle"] == DBNull.Value ? String.Empty : dr["JobTitle"].ToString();
            EmployeeID = dr["EmployeeID"] == DBNull.Value ? String.Empty : dr["EmployeeID"].ToString();
            CompletionStatus = dr["CompletionStatus"].ToString();
            SuccessStatus = dr["SuccessStatus"].ToString();
            CompletionDate = dr["CompletionDate"] == DBNull.Value ? null : (DateTime?)dr["CompletionDate"];
            Score = dr["Score"] == DBNull.Value ? decimal.MinValue : (decimal)dr["Score"];
            ParentGroup = dr["ParentArea"] == DBNull.Value ? String.Empty : dr["ParentArea"].ToString();
            Category = dr["Category"] == DBNull.Value ? String.Empty : dr["Category"].ToString();
            HireDate = dr["HireDate"] == DBNull.Value ? null : (DateTime?)dr["HireDate"];
            Manager = dr["Manager"] == DBNull.Value ? String.Empty : dr["Manager"].ToString();
        }

        private static List<CustomReport> GetCustomReport(int group1ID, DateTime hireStartDate, DateTime hireEndDate, int LMSID, int completionStatus, int successStatus, 
                                             string courseCategory, int courseID, bool mostRecentAttempts, DateTime completionStartDate, DateTime completionEndDate, 
                                             string metaTag, bool inactive, bool includeNotAttempted, List<int> courses_not_required = null, List<int> groups_not_required = null,
                                             string group1Type = "area", string group2Type = "none", int group2ID = -1, string group3Type = "none", int group3ID = -1)
        {
            // Generate custom report using the provided arguments.
            DataView dv = getCustomReportData(group1Type, group1ID, group2Type, group2ID, group3Type, group3ID, hireStartDate, hireEndDate, LMSID, completionStatus, successStatus
                                                          ,courseCategory, courseID, mostRecentAttempts, completionStartDate, completionEndDate, metaTag, inactive, includeNotAttempted);
            List<CustomReport> list = new List<CustomReport>();
            foreach(DataRow dr in dv.Table.Rows)
            {
                CustomReport singleReportRow = new CustomReport(dr);
                // Condition to check if a course should not be included in the report.
                if (courses_not_required != null && courses_not_required.Count >= 0)
                {
                    if (courses_not_required.Contains(singleReportRow.CourseID))
                    {
                        continue;
                    }
                    else
                    {
                        // Condition to check if a group should not be part of the report.
                        if (groups_not_required != null && groups_not_required.Count >= 0)
                        {
                            if (groups_not_required.Contains(singleReportRow.AreaID))
                            {
                                continue;
                            }
                            else
                            {
                                list.Add(singleReportRow);
                            }
                        }
                    }
                }
                
                
            }
            return list;
        }

        // Function to call the stored procedure to generate the custom report.
        private static DataView getCustomReportData(string group1Type, int group1ID, string group2Type, int group2ID, string group3Type, int group3ID, DateTime hireStartDate,
                                     DateTime hireEndDate, int LMSID, int completionStatus, int successStatus, string courseCategory, int courseID, bool mostRecentAttempts,
                                     DateTime completionStartDate, DateTime completionEndDate, string metaTag, bool inactive, bool includeNotAttempted)
        {
            string sql = "spGetReportData";
            var pars = new Dictionary<string, object>();

            pars.Add("@group1Type", group1Type);
            pars.Add("@group1ID", group1ID);
            pars.Add("@group2Type", group2Type);
            pars.Add("@group2ID", group2ID);
            pars.Add("@group3Type", group3Type);
            pars.Add("@group3ID", group3ID);
            pars.Add("@hireDateStart", hireStartDate);
            pars.Add("@hireDateEnd", hireEndDate);
            pars.Add("@LMSID", LMSID);
            pars.Add("@completionStatus", completionStatus);
            pars.Add("@successStatus", successStatus);
            pars.Add("@courseCategory", courseCategory);
            pars.Add("@course", courseID);
            pars.Add("@mostRecentAttempt", mostRecentAttempts);
            pars.Add("@completionDateStart", completionStartDate);
            pars.Add("@completionDateEnd", completionEndDate);
            pars.Add("@metaTag", metaTag);
            pars.Add("@inactive", inactive);
            pars.Add("@includeNotAttempted", includeNotAttempted);

            return Utility.GetDataFromQueryPortal(sql, CommandType.StoredProcedure, pars);
        }

        public static void createCustomReport(AutoReporting report, string reportFileLocation)
        {
            // Desiralize the list of argments required to run the report. 
            CustomReportModel.CustomReportArgs reportArgs = CustomReportModel.DeserializeCustomReportArgs(report.ReportArgList);

            // Start the report generation process only if the report is set to run on the current day.
            if (!CommonUtils.checkReportFrequency(reportArgs.frequency))
            {
                return;
            }
            else
            {
                // Get the client's date format for the report
                string customDateFormatString = CommonUtils.GetDateFormatFrmOrgSettings(report.OrganizationID);

                // Hire enddate & Completion end date value. 
                DateTime hireNcompleteEndDate = DateTime.Now.Date;

                // Hire start date & Completion start date values going back from the current date.
                DateTime hireStartDate = hireNcompleteEndDate.AddDays(reportArgs.hire_days_back);
                DateTime completionStartDate = hireNcompleteEndDate.AddDays(reportArgs.completion_days_back);

                int LMSID = CommonUtils.GetLMSID(report.OrganizationID);
                // Get the report using the report arguments specified in the argument list in the auto report table.
                List<CustomReport> customReport = CustomReport.GetCustomReport(reportArgs.group, hireStartDate, hireNcompleteEndDate, LMSID, reportArgs.completion_status,
                                                                               reportArgs.success_status, reportArgs.category, reportArgs.course, reportArgs.most_recent_attempts,
                                                                               completionStartDate, hireNcompleteEndDate, reportArgs.meta, reportArgs.inactive, reportArgs.include_not_attempted,
                                                                               reportArgs.courses_not_required, reportArgs.groups_not_required);
                if (customReport.Count > 0)
                {
                    string reportPath = reportFileLocation + reportArgs.email_report_args.report_name.Trim() + "_" + report.OrganizationID + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".csv";
                    writeCustomReportToCSV(customReport, reportArgs.headers, reportPath, customDateFormatString);
                    if (File.Exists(reportPath))
                    {
                        CommonUtils.sendEmail(reportArgs.email_report_args.report_email_subject, reportArgs.recipients, reportArgs.email_report_args.report_email_message,
                                              reportArgs.email_report_args.report_name, report.OrganizationID, reportPath, reportArgs.report_as_attachment);
                    }
                }
            }
            
        }

        private static void writeCustomReportToCSV(List<CustomReport> reportRows, List<string> reportArgHeaders, string reportLocation, string customDateFormat = null)
        {
            List<string> headers = reportArgHeaders;
            using (StreamWriter writer = new StreamWriter(reportLocation))
            {
                // Write the headers to the csv report.
                writer.WriteLine(string.Join(",", headers));
                foreach (CustomReport customReportRow in reportRows)
                {
                    string[] selectedColumns = new string[headers.Count];
                    for (int i = 0; i < headers.Count; i++)
                    {
                        // Based on the header traverse through the CustomReport row and add to csv.
                        string header = headers[i];
                        switch (header)
                        {
                            case "CourseName":
                                selectedColumns[i] = "\"" + customReportRow.ActivityName.Replace("\"", "\"\"") + "\"";
                                break;
                            case "UserName":
                                selectedColumns[i] = customReportRow.UserName;
                                break;
                            case "FirstName":
                                selectedColumns[i] = customReportRow.FirstName;
                                break;
                            case "LastName":
                                selectedColumns[i] = customReportRow.LastName;
                                break;
                            case "Email":
                                selectedColumns[i] = customReportRow.Email;
                                break;
                            case "Group":
                                selectedColumns[i] = customReportRow.GroupName;
                                break;
                            case "Job":
                                selectedColumns[i] = customReportRow.JobName;
                                break;
                            case "EmployeeID":
                                selectedColumns[i] = customReportRow.EmployeeID;
                                break;
                            case "CompletionStatus":
                                selectedColumns[i] = customReportRow.CompletionStatus;
                                break;
                            case "SuccessStatus":
                                selectedColumns[i] = customReportRow.SuccessStatus;
                                break;
                            case "CompletionDate":
                                selectedColumns[i] = Convert.ToDateTime(customReportRow.CompletionDate).ToString(customDateFormat);
                                break;
                            case "Score":
                                selectedColumns[i] = Decimal.Round(customReportRow.Score, 2).ToString();
                                break;
                            case "ParentGroup":
                                selectedColumns[i] = customReportRow.ParentGroup;
                                break;
                            case "Category":
                                selectedColumns[i] = customReportRow.Category;
                                break;
                            case "HireDate":
                                selectedColumns[i] = Convert.ToDateTime(customReportRow.HireDate).ToString(customDateFormat);
                                break;
                            case "Manager":
                                selectedColumns[i] = customReportRow.Manager;
                                break;

                        }
                    }
                    writer.WriteLine(string.Join(",", selectedColumns));
                }
            }
        }
    }

    
}
