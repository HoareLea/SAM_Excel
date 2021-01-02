using Microsoft.Office.Interop.Excel;

namespace SAM.Core.Excel
{
    public static partial class Query
    {
        public static bool Edited(Application Application)
        {            
            if (Application == null)
                return false;

            if (!Application.Interactive)
                return false;

            try
            {
                Application.Interactive = false;
                Application.Interactive = true;
            }
            catch
            {
                return true;
            }

            return false;
        }
    }
}