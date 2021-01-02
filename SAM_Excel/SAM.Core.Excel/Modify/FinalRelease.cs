using System;

namespace SAM.Core.Excel
{
    public static partial class Modify
    {
        public static void FinalRelease()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}


