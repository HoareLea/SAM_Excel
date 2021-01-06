using NetOffice.ExcelApi;
using System;

namespace SAM.Core.Excel
{
    public static partial class Query
    {
        public static bool Edited(this Application application)
        {
            if (!application.Interactive)
                return false;

            try
            {
                application.Interactive = false;
                application.Interactive = true;
            }
            catch
            {
                return true;
            }

            return false;
        }

        public static bool Edited()
        {
            Application application = null;
            bool result = false;
            try
            {
                application = Application.GetActiveInstance(false);
                if (application != null)
                    result = Edited(application);
            }
            catch (Exception exception)
            {
                result = false;
            }
            finally
            {
                if (application != null)
                {
                    application.Quit();
                    application.Dispose();
                }
            }

            return result;
        }
    }
}