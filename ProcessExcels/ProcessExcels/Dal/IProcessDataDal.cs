using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessExcels.Dal
{
    public interface IProcessDataDal
    {
        string BuildInsertQuery(string tablename, Dictionary<string, object> keyValues);
        void InsertRecords(string querystring);
        bool TableAvailableInDataBase(string tableName);
        void CreateTabel(string queryString);
        string BuildCreateTableQuery(string tablename, Dictionary<string, object> keyValues);
    }
}
