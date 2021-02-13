using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using SAM.Core.Grasshopper.Excel.Properties;
using System;
using System.Collections.Generic;

namespace SAM.Core.Grasshopper.Excel
{
    public class SAMExcelWrite : GH_SAMVariableOutputParameterComponent
    {
        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid => new Guid("c51ea883-e931-41c0-8e44-0be46627c51a");

        /// <summary>
        /// The latest version of this component
        /// </summary>
        public override string LatestComponentVersion => "1.0.1";

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon => Resources.SAM_Small;

        public override GH_Exposure Exposure => GH_Exposure.primary;

        /// <summary>
        /// Initializes a new instance of the SAM_point3D class.
        /// </summary>
        public SAMExcelWrite()
          : base("SAMExcel.Write", "SAMExcel.Write",
              "Write Microsoft Excel File",
              "SAM", "Excel")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override GH_SAMParam[] Inputs
        {
            get
            {
                List<GH_SAMParam> result = new List<GH_SAMParam>();
                result.Add(new GH_SAMParam(new Param_String() { Name = "_path", NickName = "_path", Description = "Path to Microsoft Excel File", Access = GH_ParamAccess.item }, ParamVisibility.Binding));
                result.Add(new GH_SAMParam(new Param_String() { Name = "_worksheetName", NickName = "_worksheetName", Description = "Worksheet Name", Access = GH_ParamAccess.item }, ParamVisibility.Binding));
                result.Add(new GH_SAMParam(new Param_GenericObject() { Name = "_values", NickName = "_values", Description = "Values", Access = GH_ParamAccess.tree }, ParamVisibility.Binding));

                Param_Integer integer = null;

                integer = new Param_Integer() { Name = "_rowIndex_", NickName = "_rowIndex_", Description = "Start Row Index", Access = GH_ParamAccess.item };
                integer.SetPersistentData(1);
                result.Add(new GH_SAMParam(integer, ParamVisibility.Binding));

                integer = new Param_Integer() { Name = "_columnIndex_", NickName = "_columnIndex_", Description = "Start Column Index", Access = GH_ParamAccess.item };
                integer.SetPersistentData(1);
                result.Add(new GH_SAMParam(integer, ParamVisibility.Binding));

                Param_Boolean param_Boolean = new Param_Boolean() { Name = "_run_", NickName = "_run_", Description = "Run", Access = GH_ParamAccess.item };
                param_Boolean.SetPersistentData(false);

                result.Add(new GH_SAMParam(param_Boolean, ParamVisibility.Binding));
                return result.ToArray();
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override GH_SAMParam[] Outputs
        {
            get
            {
                List<GH_SAMParam> result = new List<GH_SAMParam>();

                result.Add(new GH_SAMParam(new Param_Boolean() { Name = "Successful", NickName = "Successful", Description = "Successful", Access = GH_ParamAccess.item }, ParamVisibility.Binding));
                return result.ToArray();
            }
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="dataAccess">
        /// The DA object is used to retrieve from inputs and store in outputs.
        /// </param>
        protected override void SolveInstance(IGH_DataAccess dataAccess)
        {
            int index = -1;
            
            bool run = false;
            index = Params.IndexOfInputParam("_run_");
            if (index == -1 || !dataAccess.GetData(index, ref run) || !run)
                return;

            string path = null;
            index = Params.IndexOfInputParam("_path");
            if (index == -1 || !dataAccess.GetData(index, ref path) || string.IsNullOrEmpty(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid data");
                return;
            }

            if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(path)))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Directory does not exists");
                return;
            }

            string worksheetName = null;
            index = Params.IndexOfInputParam("_worksheetName");
            if (index == -1 || !dataAccess.GetData(index, ref worksheetName) || string.IsNullOrEmpty(worksheetName))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid data");
                return;
            }

            global::Grasshopper.Kernel.Data.GH_Structure<global::Grasshopper.Kernel.Types.IGH_Goo> structure;
            index = Params.IndexOfInputParam("_values");
            if (index == -1 || !dataAccess.GetDataTree(index, out structure))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid data");
                return;
            }

            int rowIndex = 1;
            index = Params.IndexOfInputParam("_rowIndex_");
            if (index != -1)
                dataAccess.GetData(index, ref rowIndex);

            int columnIndex = 1;
            index = Params.IndexOfInputParam("_columnIndex_");
            if (index != -1)
                dataAccess.GetData(index, ref columnIndex);

            object[,] values = Query.Objects(structure);
            bool result = Core.Excel.Modify.Write(path, worksheetName, values, rowIndex, columnIndex);

            index = Params.IndexOfOutputParam("Successful");
            if (index != -1)
                dataAccess.SetData(index, result);
        }
    }
}