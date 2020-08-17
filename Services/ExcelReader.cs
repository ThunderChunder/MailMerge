using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using ExcelDataReader;
using MailMerge.Contracts;

namespace MailMerge.Services
{
    public class ExcelReader: IExcelReader
    {
        public DataSet CreateDataSet()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(),"wwwroot/spreadsheets", "spreadsheet.xlsx");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;

                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = false 
                    }
                };

                return reader.AsDataSet(conf);

                // var dataSet = reader.AsDataSet(conf);
                // var dataTable = dataSet.Tables[0];

                // // Now you can get data from each sheet by its index or its "name"

                // for (var i = 0; i < dataTable.Rows.Count; i++)
                // {
                //     for (var j = 0; j < dataTable.Columns.Count; j++)
                //     {
                //         var data = dataTable.Rows[i][j];
                //         Console.WriteLine(data);
                //     }
                // }
                // return dataSet;
            }

        }
    }
}