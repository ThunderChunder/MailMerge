using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using MailMerge.Contracts;

namespace MailMerge.Services
{
    public class ExcelReader: IExcelReader
    {
        public List<List<string>> Create2DList()
        {
            List<List<string>> spreadSheet = new List<List<string>>();
            List<string> tempList;

            var path = Path.Combine(Directory.GetCurrentDirectory(),"wwwroot/spreadsheets", "spreadsheet.xlsx");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read()) //Each ROW
                    {
                        tempList = new List<string>();
                        for (int column = 0; column < reader.FieldCount; column++)
                        {
                            tempList.Add(reader.GetValue(column).ToString());
                        }
                        spreadSheet.Add(tempList);
                    }
                }
            }
            return spreadSheet;
        }
    }
}