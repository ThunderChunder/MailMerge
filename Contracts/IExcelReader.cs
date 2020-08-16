using System.Collections.Generic;

namespace MailMerge.Contracts
{
    public interface IExcelReader
    {
         public List<List<string>> Create2DList();
    }
}