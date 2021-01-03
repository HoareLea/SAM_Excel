using NetOffice.ExcelApi;

namespace SAM.Core.Excel
{
    public static partial class Create
    {
        public static DelimitedFileTable DelimitedFileTable(string path, string worksheetName, int namesIndex = 1, int headerCount = 0)
        {
            if (string.IsNullOrWhiteSpace(path) || string.IsNullOrEmpty(worksheetName))
                return null;

            if (!System.IO.File.Exists(path))
                return null;

            object[,] values = Query.Values(path, worksheetName);
            if (values == null)
                return null;

            return new DelimitedFileTable(values, namesIndex, headerCount);
        }
        
        public static DelimitedFileTable DelimitedFileTable(this Worksheet worksheet, int namesIndex = 1, int headerCount = 0)
        {
            if (worksheet == null)
                return null;

            object[,] values = worksheet.Values();
            if (values == null)
                return null;

            return new DelimitedFileTable(values, namesIndex, headerCount);
        }
    }
}