using System;
using System.Data;
using System.Data.OleDb;

namespace ProcessExcels.Bal
{
    public class ExcelOperations : IExcelOperations
    {
        private const string CONNECTION_STRING = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=1;'";
        public string[] GetAllSheetsFromExcelFile(string fileName)
        {
            using (OleDbConnection connection = new OleDbConnection())
            {
                connection.ConnectionString = string.Format(CONNECTION_STRING, fileName);
                connection.Open();
                DataTable dataTable = new DataTable();
                dataTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dataTable == null)
                {
                    return null;
                }

                string[] excelSheets = new string[dataTable.Rows.Count];
                int i = 0;
                foreach (DataRow row in dataTable.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }
                return excelSheets;
            }
        }

        public DataTable GetRecordsFromSheet(string fileName, string sheetName)
        {
            using (OleDbConnection connection = new OleDbConnection())
            {
                connection.ConnectionString = string.Format(CONNECTION_STRING, fileName);
                using (OleDbCommand command = new OleDbCommand())
                {
                    command.CommandText = string.Format("SELECT * FROM [{0}]", sheetName);
                    command.Connection = connection;
                    DataTable dataTable = new DataTable();

                    connection.Open();
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        dataTable.Load(reader);
                        return dataTable;
                    }
                }
            }
        }
    }
}
