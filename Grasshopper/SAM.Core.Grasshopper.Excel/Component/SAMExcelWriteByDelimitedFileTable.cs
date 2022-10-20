using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using SAM.Core.Grasshopper.Excel.Properties;
using System;
using System.Collections.Generic;

namespace SAM.Core.Grasshopper.Excel
{
    public class SAMExcelWriteByDelimitedFileTable : GH_SAMVariableOutputParameterComponent
    {
        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid => new Guid("67a801cd-f0ae-4098-b0a9-bf575b79ea75");

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
        /// Initializes a new instance of the SAM_point3D class.
        /// </summary>
        public SAMExcelWriteByDelimitedFileTable()
          : base("SAMExcel.WriteByDelimitedFileTable", "SAMExcel.WriteByDelimitedFileTable",
              "Write Microsoft Excel File by DelimitedFileTable",
              "SAM WIP", "Excel")
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
                result.Add(new GH_SAMParam(new GooDelimitedFileTableParam() { Name = "_delimitedFileTable", NickName = "_delimitedFileTable", Description = "DelimitedFileTable", Access = GH_ParamAccess.item }, ParamVisibility.Binding));

                Param_Integer integer = null;

                integer = new Param_Integer() { Name = "_rowIndex_", NickName = "_rowIndex_", Description = "Start Row Index", Access = GH_ParamAccess.item };
                integer.SetPersistentData(1);
                result.Add(new GH_SAMParam(integer, ParamVisibility.Binding));

                integer = new Param_Integer() { Name = "_columnIndex_", NickName = "_columnIndex_", Description = "Start Column Index", Access = GH_ParamAccess.item };
                integer.SetPersistentData(1);
                result.Add(new GH_SAMParam(integer, ParamVisibility.Binding));

                Param_String param_String = new Param_String() { Name = "_clearOption_", NickName = "_clearOption_", Description = "Clear Option", Access = GH_ParamAccess.item };
                param_String.SetPersistentData(Core.Excel.ClearOption.Data.ToString());
                result.Add(new GH_SAMParam(param_String, ParamVisibility.Binding));

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

                result.Add(new GH_SAMParam(new Param_Boolean() { Name = "successful", NickName = "successful", Description = "Successful", Access = GH_ParamAccess.item }, ParamVisibility.Binding));
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

            int rowIndex = 1;
            index = Params.IndexOfInputParam("_rowIndex_");
            if (index != -1)
                dataAccess.GetData(index, ref rowIndex);

            int columnIndex = 1;
            index = Params.IndexOfInputParam("_columnIndex_");
            if (index != -1)
                dataAccess.GetData(index, ref columnIndex);

            Core.Excel.ClearOption clearOption = Core.Excel.ClearOption.Data;
            index = Params.IndexOfInputParam("_clearOption_");
            if (index != -1)
            {
                string clearOptionString = null;
                if (dataAccess.GetData(index, ref clearOptionString))
                    Enum.TryParse(clearOptionString, out clearOption);
            }

            bool result = Core.Excel.Modify.Write(path, worksheetName, delimitedFileTable, rowIndex, columnIndex, clearOption);

            //Wait 2 sek
            System.Threading.Thread.Sleep(1000);

            index = Params.IndexOfOutputParam("successful");
            if (index != -1)
                dataAccess.SetData(index, result);
        }
    }
}