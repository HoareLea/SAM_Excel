using System.IO;
using NetOffice.ExcelApi;

namespace SAM.Core.Excel
{
    public static partial class Convert
    {
        public static string ToPdf(string excelPath)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                return null;
            }

            string pdfPath = Path.Combine(Path.GetDirectoryName(excelPath), Path.GetFileNameWithoutExtension(excelPath) + ".pdf");
            return ToPdf(excelPath, pdfPath);
        }

        public static string ToPdf(string excelPath, string pdfPath)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath) || string.IsNullOrWhiteSpace(pdfPath))
            {
                return null;
            }

            if(!Directory.Exists(Path.GetDirectoryName(pdfPath)))
            {
                return null;
            }

            Application application = new Application();
            application.DisplayAlerts = false;

            Workbook workbook = application.Workbooks.Open(excelPath);

            workbook.ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType.xlTypePDF, pdfPath);

            application.Quit();
            application.Dispose();

            return pdfPath;
        }
    }
}

