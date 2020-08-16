using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using MailMerge.Contracts;

namespace MailMerge.Processors
{
    public class ExcelReader : IExcelReader
    {
        public List<List<string>> Create2DList()
        {
            List<List<string>> spreadSheet = new List<List<string>>();
            List<string> temp;

            var path = Path.Combine(Directory.GetCurrentDirectory(),"wwwroot/spreadsheets", "f1df800b-1eef-4c8a-9a2d-fad734537178.xlsx");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read()) //Each ROW
                    {
                        temp = new List<string>();
                        for (int column = 0; column < reader.FieldCount; column++)
                        {
                            temp.Add(reader.GetValue(column).ToString());
                        }
                        spreadSheet.Add(temp);
                    }
                }
            }
            return spreadSheet;
        }
    }
}