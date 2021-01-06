using NetOffice.ExcelApi;
using System;

namespace SAM.Core.Excel
{
    public static partial class Modify
    {
        public static bool RunMacro(string path, string macroName, params object[] arguments)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;

            if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(path)))
                return false;

            if (!System.IO.File.Exists(path))
                return false;

            Application application = null;
            bool result = false;
            try
            {
                application = new Application(true);
                application.DisplayAlerts = false;
                application.Visible = false;
                application.ScreenUpdating = false;

                Workbook workbook = application.Workbooks.Open(path);
                if (workbook != null)
                    result = RunMacro(application, macroName, arguments);
            }
            catch (Exception exception)
            {
                result = false;
            }
            finally
            {
                if (application != null)
                {
                    application.ScreenUpdating = true;

                    application.Quit();
                    application.Dispose();
                }
            }

            return result;
        }

        public static bool RunMacro(Application application, string macroName, params object[] arguments)
        {
            if (application == null)
                return false;
            if (string.IsNullOrEmpty(macroName))
                return false;

            try
            {
                if (arguments == null || arguments.Length == 0)
                    application.Run(macroName);
                else if (arguments.Length == 1)
                    application.Run(macroName, arguments[0]);
                else if (arguments.Length == 2)
                    application.Run(macroName, arguments[0], arguments[1]);

                return true;
            }
            catch (Exception exception)
            {
                return false;
            }

        }
    }
}


