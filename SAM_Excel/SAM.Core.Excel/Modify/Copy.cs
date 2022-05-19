using NetOffice.ExcelApi;
using System.Collections.Generic;
using System.Linq;

namespace SAM.Core.Excel
{
    public static partial class Modify
    {
        public static Worksheet Copy(Worksheet worksheet, string name = null, int index = int.MaxValue)
        {
            Workbook workbook = worksheet?.Workbook();
            if(workbook == null)
            {
                return null;
            }

            Sheets sheets = workbook.Sheets;
            if(sheets == null || sheets.Count == 0)
            {
                return null;
            }

            List<string> worksheetNames_Before = workbook.WorksheetNames();

            if(index >= sheets.Count)
            {
                worksheet.Copy(null, sheets[sheets.Count]);
            }
            else if(index < 0)
            {
                worksheet.Copy(sheets[1]);
            }
            else
            {
                worksheet.Copy(sheets[index + 1]);
            }

            List<string> worksheetNames_After = workbook.WorksheetNames();

            string worksheetName = worksheetNames_After.FindAll(x => !worksheetNames_Before.Contains(x)).FirstOrDefault();
            if(string.IsNullOrWhiteSpace(worksheetName))
            {
                return null;
            }

            Worksheet result = workbook.Worksheets[worksheetName] as Worksheet;

            if(result != null && !string.IsNullOrEmpty(name))
            {
                result.Name = name;
            }

            return result;
        }
    }
}


