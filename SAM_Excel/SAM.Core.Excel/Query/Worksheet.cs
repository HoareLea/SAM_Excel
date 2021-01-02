using NetOffice.ExcelApi;

namespace SAM.Core.Excel
{
    public static partial class Query
    {
        public static Worksheet Worksheet(this Workbook workbook, string name)
        {
            if (workbook == null || workbook.Worksheets == null|| string.IsNullOrEmpty(name))
                return null;

            int count = workbook.Worksheets.Count;

            for(int i=0; i < count; i++)
            {
                Worksheet worksheet = workbook.Worksheets[i + 1] as Worksheet;
                if (name.Equals(worksheet?.Name))
                    return worksheet;
            }

            return null;
        }
    }
}