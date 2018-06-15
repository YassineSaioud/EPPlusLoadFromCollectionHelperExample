using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace RLD.Extranet.Helpers
{
    public static class ExportHelper
    {
        public static FileInfo BuildWorkbookFromCollection<T>(string title, IDictionary<string, string> columns, IEnumerable<T> collection) where T : class
        {
            using (var pck = new ExcelPackage())
            {
                pck.Workbook.Properties.Title = title;

                var ws = pck.Workbook.Worksheets.Add(title);

                // Building Header
                var rowIndex = 1;
                var columnIndex = 1;
                foreach (var header in columns)
                {
                    ws.Cells[rowIndex, columnIndex].Value = header.Key;
                    columnIndex += 1;
                }

                // Building rows
                rowIndex = 2;
                foreach (var item in collection)
                {
                    columnIndex = 1;
                    foreach (var column in columns)
                    {
                        var columnValue = GetPropertyValue(item, column.Value);
                        if (columnValue is DateTime)
                            ws.Cells[rowIndex, columnIndex].Style.Numberformat.Format = "dd/MM/yyyy";

                        ws.Cells[rowIndex, columnIndex].Value = columnValue;
                        ws.Cells[rowIndex, columnIndex].AutoFitColumns();

                        columnIndex += 1;
                    }
                    rowIndex += 1;
                }

                using (MemoryStream memoryStream = new MemoryStream())
                {
                    pck.SaveAs(memoryStream);
                    memoryStream.Position = 0;
                    return new FileInfo
                    {
                        Guid = Guid.NewGuid().ToString(),
                        Name = $"{title}-{DateTime.Now.ToShortDateString()}.xlsx",
                        Contents = memoryStream.ToArray()
                    };
                }

            }
        }

        #region Private Methode
        static object GetPropertyValue(object src, string propName)
        {
            if (src == null) throw new ArgumentException("Value cannot be null.", nameof(src));
            if (propName == null) throw new ArgumentException("Value cannot be null.", nameof(propName));

            object value = null;

            if (propName.Contains("."))
            {
                var nestedProperty = propName.Split(new char[] { '.' }, 2);
                value = GetPropertyValue(GetPropertyValue(src, nestedProperty[0]), nestedProperty[1]);
            }
            else
            {
                var property = src.GetType().GetProperty(propName);
                value = property?.GetValue(src, null);
            }

            return value;
        }
        #endregion

    }

    public class FileInfo
    {
        public string Guid { get; set; }
        public string Name { get; set; }
        public Byte[] Contents { get; set; }
    }

}

