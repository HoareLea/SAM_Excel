using Grasshopper;
using Grasshopper.Kernel;
using SAM.Core.Grasshopper.Excel.Properties;

namespace SAM.Core.Grasshopper.Excel
{
    public class AssemblyPriority : GH_AssemblyPriority
    {
        public override GH_LoadingInstruction PriorityLoad()
        {
            Instances.ComponentServer.AddCategoryIcon("SAM", Resources.SAM_Small);
            Instances.ComponentServer.AddCategorySymbolName("SAM", 'S');
            return GH_LoadingInstruction.Proceed;
        }
    }
}