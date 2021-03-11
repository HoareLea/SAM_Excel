using SAM.Core.Grasshopper.Excel.Properties;
using System;
using Grasshopper.Kernel;

namespace SAM.Core.Grasshopper.Excel
{
    public class SAMExcelClearOption : GH_SAMEnumComponent<Core.Excel.ClearOption>
    {
        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid => new Guid("a42f5b20-3d28-4067-8fb8-1e1b309895b7");

        /// <summary>
        /// The latest version of this component
        /// </summary>
        public override string LatestComponentVersion => "1.0.0";

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon => Resources.SAM_Excel;

        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        /// <summary>
        /// Clear Option Enum Component
        /// </summary>
        public SAMExcelClearOption()
          : base("SAMExcel.ClearOption", "SAMExcel.ClearOption",
              "Clear Option",
              "SAM", "Excel")
        {
        }
    }
}