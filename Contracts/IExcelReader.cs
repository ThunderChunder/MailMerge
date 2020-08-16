using System.Collections.Generic;

namespace MailMerge.Contracts
{
    public interface IExcelReader
    {
         public string[,] Create2DArray();
    }
}