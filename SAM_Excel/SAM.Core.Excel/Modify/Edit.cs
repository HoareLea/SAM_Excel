using NetOffice.ExcelApi;
using System;

namespace SAM.Core.Excel
{
    public static partial class Modify
    {
        public static bool Edit(string path, string worksheetName, Func<Worksheet, bool> func)
        {
            if (string.IsNullOrWhiteSpace(path) || func == null)
            {
                return false;
            }

            Func<Workbook, bool> func_Workbook = new Func<Workbook, bool>((Workbook workbook) =>
            {
                Worksheet worksheet = workbook.Worksheet(worksheetName);
                if(worksheet == null)
                {
                    return false;
                }

                return func.Invoke(worksheet);
            });

            return Edit(path, func_Workbook);
        }

        public static bool Edit(string path, Func<Workbook, bool> func)
        {
            if(string.IsNullOrWhiteSpace(path) || func == null)
            {
                return false;
            }

            Func<Application, bool> func_Application = new Func<Application, bool>((Application application) =>
            {
                Workbook workbook = null;
                if (System.IO.File.Exists(path))
                    workbook = application.Workbooks.Open(path);
                else
                    workbook = application.Workbooks.Add();

                bool result = func.Invoke(workbook);
                if(result)
                {
                    workbook.SaveAs(path);
                }

                workbook.Close(false);
                return result;
            });

            return Edit(func_Application);
        }

        public static bool Edit(Func<Application, bool> func)
        {
            if (func == null)
            {
                return false;
            }

            bool result = false;

            Application application = null;

            bool screenUpdating = false;
            bool displayStatusBar = false;
            bool enableEvents = false;

            try
            {
                application = new Application();
                application.DisplayAlerts = false;
                application.Visible = false;

                screenUpdating = application.ScreenUpdating;
                displayStatusBar = application.DisplayStatusBar;
                enableEvents = application.EnableEvents;

                application.ScreenUpdating = false;
                application.DisplayStatusBar = false;
                application.EnableEvents = false;

                result = func.Invoke(application);
            }
            catch (Exception exception)
            {

            }
            finally
            {
                if (application != null)
                {
                    application.ScreenUpdating = screenUpdating;
                    application.DisplayStatusBar = displayStatusBar;
                    application.EnableEvents = enableEvents;

                    application.Quit();
                    application.Dispose();
                }
            }

            return result;
        }
    }
}
