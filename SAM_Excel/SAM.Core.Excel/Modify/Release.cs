using System.Runtime.InteropServices;

namespace SAM.Core.Excel
{
    public static partial class Modify
    {
        public static void Release(object Object)
        {
            if (Object != null && Marshal.IsComObject(Object))
                Marshal.ReleaseComObject(Object);
        }
    }
}


