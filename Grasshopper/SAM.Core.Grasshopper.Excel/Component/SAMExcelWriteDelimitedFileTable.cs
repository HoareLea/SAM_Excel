using Grasshopper;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using SAM.Core.Grasshopper.Excel.Properties;
using System;
using System.Collections.Generic;

namespace SAM.Core.Grasshopper.Excel
{
    public class SAMExcelWriteDelimitedFileTable : GH_SAMVariableOutputParameterComponent
    {
        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid => new Guid("9a48d83b-b9ca-4923-a8aa-0c65c18ce638");

        /// <summary>
        /// The latest version of this component
        /// </summary>
        public override string LatestComponentVersion => "1.0.0";

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon => Resources.SAM_Small;

        public override GH_Exposure Exposure => GH_Exposure.primary;

        /// <summary>
        /// Initializes a new instance of the SAM_point3D class.
        /// </summary>
        public SAMExcelWriteDelimitedFileTable()
          : base("SAMExcel.WriteDelimitedFileTable", "SAMExcel.WriteDelimitedFileTable",
              "Write DelimitedFileTable to Microsoft Excel File",
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
                result.Add(new GH_SAMParam(new GooDelimitedFileTableParam() { Name = "_delimitedFileTable", NickName = "_delimitedFileTable", Description = "SAM DelimitedFileTable", Access = GH_ParamAccess.item }, ParamVisibility.Binding));

                Param_Integer param_Integer = null;
                Param_Boolean param_Boolean = null;

                param_Integer = new Param_Integer() { Name = "_headerIndex_", NickName = "_headerIndex_", Description = "Header Index", Access = GH_ParamAccess.item};
                param_Integer.SetPersistentData(1);
                result.Add(new GH_SAMParam(param_Integer, ParamVisibility.Binding));

                param_Integer = new Param_Integer() { Name = "_headerCount_", NickName = "_headerCount_", Description = "Header Count", Access = GH_ParamAccess.item };
                param_Integer.SetPersistentData(0);
                result.Add(new GH_SAMParam(param_Integer, ParamVisibility.Binding));

                param_Boolean = new Param_Boolean() { Name = "_clear_", NickName = "_clear_", Description = "Clear before data insert", Access = GH_ParamAccess.item };
                param_Boolean.SetPersistentData(false);
                result.Add(new GH_SAMParam(param_Boolean, ParamVisibility.Binding));

                param_Boolean = new Param_Boolean() { Name = "_run_", NickName = "_run_", Description = "Run", Access = GH_ParamAccess.item };
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

                result.Add(new GH_SAMParam(new Param_Boolean() { Name = "Successful", NickName = "Successful", Description = "Successful", Access = GH_ParamAccess.tree }, ParamVisibility.Binding));
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
            {
                dataAccess.SetData(0, false);
                return;
            }

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

            DelimitedFileTable delimitedFileTable = null;
            index = Params.IndexOfInputParam("_delimitedFileTable");
            if (index == -1 || !dataAccess.GetData(index, ref delimitedFileTable))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid data");
                return;
            }

            int headerIndex = 1;
            index = Params.IndexOfInputParam("_headerIndex_");
            if (index != -1)
                dataAccess.GetData(index, ref headerIndex);

            int headerCount = 1;
            index = Params.IndexOfInputParam("_headerCount_");
            if (index != -1)
                dataAccess.GetData(index, ref headerCount);

            bool clear = false;
            index = Params.IndexOfInputParam("_clear_");
            if (index != -1)
                dataAccess.GetData(index, ref clear);

            bool result = Core.Excel.Modify.Update(path, worksheetName, delimitedFileTable, headerIndex, headerCount, clear);

            //Wait 2 sek
            System.Threading.Thread.Sleep(1000);
            index = Params.IndexOfOutputParam("Successful");
            if (index != -1)
                dataAccess.SetData(index, result);
        }
    }
}