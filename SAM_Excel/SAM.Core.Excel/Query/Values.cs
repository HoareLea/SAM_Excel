using NetOffice.ExcelApi;
using System;

namespace SAM.Core.Excel
{
    public static partial class Query
    {
        public static object[,] Values(string path, string worksheetName)
        {
            if (string.IsNullOrWhiteSpace(path) || string.IsNullOrEmpty(worksheetName))
                return null;

            if (!System.IO.File.Exists(path))
                return null;

            Application application = null;
            object[,] result = null;

            try
            {
                application = new Application();
                application.DisplayAlerts = false;

                Workbook workbook = application.Workbooks.Open(path);
                Worksheet worksheet = workbook.Worksheet(worksheetName);
                if (worksheet != null)
                    result = (object[,])worksheet.UsedRange.Value;
            }
            catch(Exception exception)
            {

            }
            finally
            {
                if(application != null)
                {
                    application.Quit();
                    application.Dispose();
                }
            }

            return result;
        }
    }
}