using NetOffice.ExcelApi;

namespace SAM.Core.Excel
{
    public static partial class Query
    {
        public static Workbook Workbook(this Worksheet worksheet)
        {
            Sheets sheets = worksheet.ParentObject as Sheets;
            if(sheets == null)
            {
                return null;
            }

            return sheets.ParentObject as Workbook;
        }
    }
}