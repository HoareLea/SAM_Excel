using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using SAM.Core.Grasshopper.Excel.Properties;
using System;
using System.Collections.Generic;

namespace SAM.Core.Grasshopper.Excel
{
    public class SAMExcelReadDelimitedFileTable : GH_SAMVariableOutputParameterComponent
    {
        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid => new Guid("41b0a87a-99aa-47b1-8ac2-e90adcc6d787");

        /// <summary>
        /// The latest version of this component
        /// </summary>
        public override string LatestComponentVersion => "1.0.0";

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon => Resources.SAM_Excel;

        public override GH_Exposure Exposure => GH_Exposure.primary;

        /// <summary>
        /// Initializes a new instance of the SAM_point3D class.
        /// </summary>
        public SAMExcelReadDelimitedFileTable()
          : base("SAMExcel.ReadDelimitedFileTable", "SAMExcel.ReadDelimitedFileTable",
              "Read Microsoft Excel File as DelimitedFileTable",
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

                Param_Integer param_Integer = null;

                param_Integer = new Param_Integer() { Name = "_namesIndex_", NickName = "_namesIndex_", Description = "Row index of column names", Access = GH_ParamAccess.item };
                param_Integer.SetPersistentData(1);
                result.Add(new GH_SAMParam(param_Integer, ParamVisibility.Binding));

                param_Integer = new Param_Integer() { Name = "_headerCount_", NickName = "_headerCount_", Description = "Row count with additional data", Access = GH_ParamAccess.item };
                param_Integer.SetPersistentData(0);
                result.Add(new GH_SAMParam(param_Integer, ParamVisibility.Binding));

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

                result.Add(new GH_SAMParam(new GooDelimitedFileTableParam() { Name = "DelimitedFileTable", NickName = "DelimitedFileTable", Description = "SAM DelimitedFileTable", Access = GH_ParamAccess.item }, ParamVisibility.Binding));
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
            int index;

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

            if (!System.IO.File.Exists(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "File does not exists");
                return;
            }

            string worksheetName = null;
            index = Params.IndexOfInputParam("_worksheetName");
            if (index == -1 || !dataAccess.GetData(index, ref worksheetName) || string.IsNullOrEmpty(worksheetName))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid data");
                return;
            }

            object[,] values = Core.Excel.Query.Values(path, worksheetName);

            if (values == null)
                return;

            int names_Index = 1;
            index = Params.IndexOfInputParam("_namesIndex_");
            if (index != -1)
                dataAccess.GetData(index, ref names_Index);

            int headerCount = 0;
            index = Params.IndexOfInputParam("_headerCount_");
            if (index != -1)
                dataAccess.GetData(index, ref headerCount);

            DelimitedFileTable delimitedFileTable = new DelimitedFileTable(values, names_Index, headerCount);

            index = Params.IndexOfOutputParam("DelimitedFileTable");
            if (index != -1)
                dataAccess.SetData(index, new GooDelimitedFileTable(delimitedFileTable));
        }
    }
}