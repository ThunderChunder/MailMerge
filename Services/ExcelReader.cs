using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using MailMerge.Contracts;

namespace MailMerge.Services
{
    public class ExcelReader: IExcelReader
    {
        public string[,] Create2DArray()
        {
            string[,] spreadSheet;
            string[] tempList;
            int row = 0;
            var path = Path.Combine(Directory.GetCurrentDirectory(),"wwwroot/spreadsheets", "spreadsheet.xlsx");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    spreadSheet = new string[reader.FieldCount, reader.RowCount];
                    tempList = new string[reader.FieldCount];
                    do
                    {
                        while (reader.Read()) //Each ROW
                        {
                            for(int column = 0; column < reader.FieldCount; column++)
                            {
                                try
                                {
                                    spreadSheet[row, column] = reader.GetString(column);
                                }
                                catch(NullReferenceException e){Console.WriteLine(e);}
                                
                            }
                            row++;
                        }
                    } while (reader.NextResult());
                }
            }
            return spreadSheet;
        }
    }
}