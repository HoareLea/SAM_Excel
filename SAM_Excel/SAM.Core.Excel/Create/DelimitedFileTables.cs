using NetOffice.ExcelApi;
using System.Collections.Generic;

namespace SAM.Core.Excel
{
    public static partial class Create
    {
        public static Dictionary<string, DelimitedFileTable> DelimitedFileTables(this Workbook workbook, int namesIndex = 1, int headerCount = 0)
        {
            if (workbook == null)
                return null;

            int count = workbook.Worksheets.Count;

            Dictionary<string, DelimitedFileTable>  result = new Dictionary<string, DelimitedFileTable>();
            for (int i = 0; i < count; i++)
            {
                Worksheet worksheet = workbook.Worksheets[i + 1] as Worksheet;
                if (worksheet == null)
                    continue;

                result[worksheet.Name] = DelimitedFileTable(worksheet, namesIndex, headerCount);
            }

            return result;
        }
    }
}