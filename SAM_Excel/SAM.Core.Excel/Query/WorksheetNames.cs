using NetOffice.ExcelApi;
using System;
using System.Collections.Generic;

namespace SAM.Core.Excel
{
    public static partial class Query
    {
        public static List<string> WorksheetNames(this Workbook workbook)
        {
            if (workbook == null)
                return null;

            Sheets sheets = workbook.Worksheets;
            if (sheets == null)
                return null;

            List<string> result = new List<string>();

            int count = sheets.Count;

            for (int i = 0; i < count; i++)
            {
                Worksheet worksheet = sheets[i + 1] as Worksheet;
                result.Add(worksheet?.Name);
            }

            return result;
        }

        public static List<string> WorksheetNames(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return null;

            if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(path)))
                return null;

            if (!System.IO.File.Exists(path))
                return null;

            Application application = null;
            List<string> result = new List<string>();
            try
            {
                application = new Application();
                application.DisplayAlerts = false;

                Workbook workbook = application.Workbooks.Open(path, 2, true);
                if (workbook != null)
                {
                    result = WorksheetNames(workbook);
                    workbook.Close(false);
                }
                    
            }
            catch (Exception exception)
            {
                result = null;
            }
            finally
            {
                if (application != null)
                {
                    application.Quit();
                    application.Dispose();
                }
            }

            return result;
        }
    }
}