using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using SAM.Core.Grasshopper.Excel.Properties;
using System;
using System.Collections.Generic;

namespace SAM.Core.Grasshopper.Excel
{
    public class SAMExcelRunMacro : GH_SAMVariableOutputParameterComponent
    {
        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid => new Guid("b9b483f4-9d6d-4fa7-8938-d54d477f35c8");

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
        public SAMExcelRunMacro()
          : base("SAMExcel.RunMacro", "SAMExcel.RunMacro",
              "Run MAcro on Microsoft Excel File",
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
                result.Add(new GH_SAMParam(new Param_String() { Name = "_macroName", NickName = "_macroName", Description = "Macro Name", Access = GH_ParamAccess.item }, ParamVisibility.Binding));

                Param_Boolean param_Boolean = null;

                param_Boolean = new Param_Boolean() { Name = "_save_", NickName = "_save_", Description = "Save after macro run", Access = GH_ParamAccess.item };
                param_Boolean.SetPersistentData(true);
                result.Add(new GH_SAMParam(param_Boolean, ParamVisibility.Binding));

                result.Add(new GH_SAMParam(new Param_GenericObject() { Name = "parameter1_", NickName = "parameter1_", Description = "Macro input parameter 1", Access = GH_ParamAccess.item, Optional = true }, ParamVisibility.Voluntary));
                result.Add(new GH_SAMParam(new Param_GenericObject() { Name = "parameter2_", NickName = "parameter2_", Description = "Macro input parameter 2", Access = GH_ParamAccess.item, Optional = true }, ParamVisibility.Voluntary));
                result.Add(new GH_SAMParam(new Param_GenericObject() { Name = "parameter3_", NickName = "parameter3_", Description = "Macro input parameter 3", Access = GH_ParamAccess.item, Optional = true }, ParamVisibility.Voluntary));
                result.Add(new GH_SAMParam(new Param_GenericObject() { Name = "parameter4_", NickName = "parameter4_", Description = "Macro input parameter 4", Access = GH_ParamAccess.item, Optional = true }, ParamVisibility.Voluntary));
                result.Add(new GH_SAMParam(new Param_GenericObject() { Name = "parameter5_", NickName = "parameter5_", Description = "Macro input parameter 5", Access = GH_ParamAccess.item, Optional = true }, ParamVisibility.Voluntary));
                result.Add(new GH_SAMParam(new Param_GenericObject() { Name = "parameter6_", NickName = "parameter6_", Description = "Macro input parameter 6", Access = GH_ParamAccess.item, Optional = true }, ParamVisibility.Voluntary));
                result.Add(new GH_SAMParam(new Param_GenericObject() { Name = "parameter7_", NickName = "parameter7_", Description = "Macro input parameter 7", Access = GH_ParamAccess.item, Optional = true }, ParamVisibility.Voluntary));
                result.Add(new GH_SAMParam(new Param_GenericObject() { Name = "parameter8_", NickName = "parameter8_", Description = "Macro input parameter 8", Access = GH_ParamAccess.item, Optional = true }, ParamVisibility.Voluntary));
                result.Add(new GH_SAMParam(new Param_GenericObject() { Name = "parameter9_", NickName = "parameter9_", Description = "Macro input parameter 9", Access = GH_ParamAccess.item, Optional = true }, ParamVisibility.Voluntary));
                result.Add(new GH_SAMParam(new Param_GenericObject() { Name = "parameter10_", NickName = "parameter10_", Description = "Macro input parameter 10", Access = GH_ParamAccess.item, Optional = true }, ParamVisibility.Voluntary));

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

                result.Add(new GH_SAMParam(new Param_Boolean() { Name = "Succeed", NickName = "Succeed", Description = "Succeded", Access = GH_ParamAccess.item }, ParamVisibility.Binding));
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
            if (index ==-1 || !dataAccess.GetData(index, ref path) || string.IsNullOrEmpty(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid data");
                return;
            }

            if (!System.IO.File.Exists(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "File does not exists");
                return;
            }

            string macroName = null;
            index = Params.IndexOfInputParam("_macroName");
            if (!dataAccess.GetData(index, ref macroName) || string.IsNullOrEmpty(macroName))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid data");
                return;
            }

            bool save = true;
            index = Params.IndexOfInputParam("_save_");
            if (index != -1)
                dataAccess.GetData(index, ref save);

            List<object> parameters = new List<object>();
            for (int i = 1; i <= 10; i++)
            {
                index = Params.IndexOfInputParam(string.Format("parameter{0}_", i));
                if (index == -1)
                    continue;

                object parameter = null;
                if (!dataAccess.GetData(index, ref parameter))
                    continue;

                if (parameter is IGH_Goo)
                    parameter = (parameter as dynamic).Value;

                parameters.Add(parameter);
            }

            bool result = Core.Excel.Modify.TryRunMacro(path, save, macroName, parameters.ToArray());

            //Wait 2 sek
            System.Threading.Thread.Sleep(1000);

            index = Params.IndexOfOutputParam("Succeed");
            if (index != -1)
                dataAccess.SetData(index, result);
        }
    }
}