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
            List<string> tempList = new List<string>();;

            var path = Path.Combine(Directory.GetCurrentDirectory(),"wwwroot/spreadsheets", "spreadsheet.xlsx");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read()) //Each ROW
                        {
                            for(int column = 0; column < reader.FieldCount; column++)
                            {
                                tempList.Clear();
                                try
                                {
                                    tempList.Add(reader.GetString(column));
                                }
                                catch(NullReferenceException e){Console.WriteLine(e);}
                                
                            }
                            spreadSheet.Add(tempList);
                        }
                    } while (reader.NextResult());
                }
            }
            return spreadSheet;
        }
    }
}