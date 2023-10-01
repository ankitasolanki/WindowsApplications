using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace ProcessExcels.Bal
{
    public interface IExcelOperations
    {
        string[] GetAllSheetsFromExcelFile(string fileName);
        DataTable GetRecordsFromSheet(string fileName, string sheetName);
    }
}
