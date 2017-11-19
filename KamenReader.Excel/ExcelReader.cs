using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SpreadsheetLight;

namespace KamenReader.Excel
{
    public sealed class ExcelReader : IFileReader
    {
        public FileReaderResult Read(String fullFilepath, IList<FileReaderMap> maps, Boolean firstRowAreTitles = true)
        {
            return Read(fullFilepath, maps, "", firstRowAreTitles);
        }

        public FileReaderResult Read(String fullFilepath, IList<FileReaderMap> maps, String sheetName = "", Boolean firstRowAreTitles = true)
        {
            if (String.IsNullOrEmpty(fullFilepath))
                throw new ArgumentException("fullFilepath");
            if (!File.Exists(fullFilepath))
                throw new InvalidOperationException("File's not found.");
            if (maps == null)
                throw new ArgumentException("maps");
            if (!maps.Any())
                throw new InvalidOperationException("File Reader Map must be supplied.");

            var result = new FileReaderResult();
            using (var doc = new SLDocument(fullFilepath))
            {
                if (String.IsNullOrEmpty(sheetName))
                    sheetName = doc.GetCurrentWorksheetName();

                doc.SelectWorksheet(sheetName);
                var stats = doc.GetWorksheetStatistics();
                for (int row = 1; row <= stats.EndRowIndex; row++)
                {
                    for (int col = 1; col <= stats.EndColumnIndex; col++)
                    {
                        String data = doc.GetCellValueAsString(row, col);
                        String cleaned = String.IsNullOrEmpty(data) ? data : data.Trim();
                        if (firstRowAreTitles && row == 1)
                            result.Titles.Add(cleaned);
                        else
                            result.Data.Add(new GridData { Row = row, Column = col, CellValue = cleaned });
                    }
                }
            }

            return result;
        }
    }
}