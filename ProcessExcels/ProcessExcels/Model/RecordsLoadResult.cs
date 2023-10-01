using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessExcels.Model
{
    public class RecordsLoadResult
    {
        public string TabPageName { get; set; }
        public DataTable DataTable { get; set; }
    }
}
