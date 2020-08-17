using System.Collections.Generic;
using System.Data;
using ExcelDataReader;

namespace MailMerge.Contracts
{
    public interface IExcelReader
    {
        public  DataSet CreateDataSet();
    }
}