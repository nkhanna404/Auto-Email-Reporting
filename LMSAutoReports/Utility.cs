using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace LMSAutoReports
{
    static class Utility
    {
        private static string _connectionStringPortal
        {
            get
            {
                return System.Configuration.ConfigurationManager.ConnectionStrings["LMSDBConnectPortal"].ToString();
            }
        }

        private static string _connectionStringGadget
        {
            get
            {
                return System.Configuration.ConfigurationManager.ConnectionStrings["LMSDBConnectGadget"].ToString();
            }
        }

        public static DataView GetDataFromQueryPortal(string query, CommandType command, Dictionary<string, object> parameters = null)
        {
            var ds = new DataSet();
            var da = new SqlDataAdapter();
            da.SelectCommand = new SqlCommand();
            da.SelectCommand.CommandTimeout = 1800;
            da.SelectCommand.CommandType = command;

            using (var conn = new SqlConnection(_connectionStringPortal))
            {
                da.SelectCommand.Connection = conn;
                da.SelectCommand.CommandText = query;

                if (parameters != null)
                {
                    foreach (string key in parameters.Keys)
                    {
                        da.SelectCommand.Parameters.AddWithValue(key, parameters[key]);
                    }
                }

                da.Fill(ds);
            }
            if (ds.Tables.Count > 0)
            {
                return ds.Tables[0].DefaultView;
            }

            return null;
        }

        public static DataView GetDataFromQueryGadget(string query, CommandType command, Dictionary<string, object> parameters = null)
        {
            var ds = new DataSet();
            var da = new SqlDataAdapter();
            da.SelectCommand = new SqlCommand();
            da.SelectCommand.CommandTimeout = 1800;
            da.SelectCommand.CommandType = command;

            using (var conn = new SqlConnection(_connectionStringGadget))
            {
                da.SelectCommand.Connection = conn;
                da.SelectCommand.CommandText = query;

                if (parameters != null)
                {
                    foreach (string key in parameters.Keys)
                    {
                        da.SelectCommand.Parameters.AddWithValue(key, parameters[key]);
                    }
                }

                da.Fill(ds);
            }
            if (ds.Tables.Count > 0)
            {
                return ds.Tables[0].DefaultView;
            }

            return null;
        }
    }
}

