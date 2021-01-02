using NetOffice.ExcelApi;
using System.Collections.Generic;

namespace SAM.Core.Excel
{
    public static partial class Query
    {
        public static List<string> WorksheetNames(this Workbook workbook)
        {
            if (workbook == null || workbook.Worksheets == null)
                return null;

            List<string> result = new List<string>();

            int count = workbook.Worksheets.Count;

            for(int i=0; i < count; i++)
            {
                Worksheet worksheet = workbook.Worksheets[i + 1] as Worksheet;
                result.Add(worksheet?.Name);
            }

            return result;
        }
    }
}