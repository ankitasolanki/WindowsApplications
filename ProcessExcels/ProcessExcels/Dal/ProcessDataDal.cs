using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Text;

namespace ProcessExcels.Dal
{
    public class ProcessDataDal : IProcessDataDal
    {
        string connectionstring = "Server=DESKTOP-PIN7UV7; Database=ProcessExcelDB; integrated security=true;";

        public string BuildInsertQuery(string tableName, Dictionary<string, object> keyValues)
        {
            StringBuilder queryBuilder = new StringBuilder();
            queryBuilder.Append($"INSERT INTO {tableName} ( ");

            int ind = 0;

            foreach (var columnName in keyValues.Keys)
            {
                queryBuilder.Append($"{columnName}");
                if (ind < keyValues.Count - 1)
                {
                    queryBuilder.Append(",");
                }
                ind++;
            }

            queryBuilder.Append(" ) VALUES ( ");

            ind = 0;

            foreach (var columnName in keyValues.Keys)
            {
                //queryBuilder.Append($"@{columnName}");
                queryBuilder.Append("'");
                queryBuilder.Append($"{keyValues[columnName]}");
                queryBuilder.Append("'");
                if (ind < keyValues.Count - 1)
                {
                    queryBuilder.Append(",");
                }
                ind++;
            }

            queryBuilder.Append(");");
            return queryBuilder.ToString();
        }

        public string BuildCreateTableQuery(string tablename, Dictionary<string, object> keyValues)
        {
            StringBuilder queryBuilder = new StringBuilder();
            queryBuilder.Append($"CREATE TABLE {tablename} (");

            int index = 0;
            foreach (var columnName in keyValues.Keys)
            {
                queryBuilder.Append($"{columnName} nvarchar(50)");
                if (index < keyValues.Count - 1)
                {
                    queryBuilder.Append(",");
                }
                index++;
            }
            queryBuilder.Append(")");
            return queryBuilder.ToString();
        }

        public void InsertRecords(string querystring)
        {
            using (SqlConnection conn = new SqlConnection(connectionstring))
            {
                conn.Open();
                using (SqlCommand command = new SqlCommand(querystring, conn))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public bool TableAvailableInDataBase(string tableName)
        {
            SqlConnection connection = new SqlConnection(connectionstring);
            connection.Open();
            string tblcheckquery = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @TableName";
            using (SqlCommand command = new SqlCommand(tblcheckquery, connection))
            {
                command.Parameters.AddWithValue("@TableName", tableName);
                int tablecount = Convert.ToInt32(command.ExecuteScalar());
                if (tablecount > 0) return true;
            }
            return false;
        }

        public void CreateTabel(string queryString)
        {
            using (SqlConnection connection = new SqlConnection(connectionstring))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(queryString, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }
    }
}
