using System;
using System.Data;
using System.Collections.Generic;

namespace LMSAutoReports
{
    public class AutoReporting
    {
        #region Properties
        public CommonUtils.eReportType ReportType { get; set; }
        public int OrganizationID { get; set; }
        public string ReportArgList { get; set; }
        #endregion

        #region Constructors
        public AutoReporting() { }

        public AutoReporting(DataRow dr)
        {
            SetValues(dr);
        }
        #endregion

        #region Public Methods
        public void SetValues(DataRow dr)
        {
            ReportType = (CommonUtils.eReportType)dr["ReportType"];
            OrganizationID = (int)dr["OrganizationID"];
            ReportArgList = dr["ReportArgList"].ToString();
        }

        public static List<AutoReporting> GetAllReports()
        {
            List<AutoReporting> list = new List<AutoReporting>();
            DataView dv = getAutoReports();
            foreach(DataRow dr in dv.Table.Rows)
            {
                AutoReporting autoReport = new AutoReporting(dr);
                list.Add(autoReport);
            }
            return list;
        }

        // Function to get all the autoreports from the database.
        private static DataView getAutoReports()
        {
            string sql = "SELECT * FROM AutoReporting";
            DataView dv = Utility.GetDataFromQueryPortal(sql, CommandType.Text);
            return dv;
        }
        #endregion
    }
}
