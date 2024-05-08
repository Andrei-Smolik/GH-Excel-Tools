using System;
using System.Collections.Generic;
using System.Drawing;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Rhino.Geometry;

namespace GH_Excel_Tools
{
    public class ExcelModify : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the ExcelModify class.
        /// </summary>
        public ExcelModify()
          : base("Excel Modify", "eMod",
              "Adds cell properties",
              "X", "Data")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("CELLS", "CELLS", "", GH_ParamAccess.list);
            pManager.AddColourParameter("COLOUR", "", "", GH_ParamAccess.item, Color.Azure);
            pManager.AddBooleanParameter("BORDER", "", "", GH_ParamAccess.item, false);
            pManager.AddIntegerParameter("FONT SIZE", "", "", GH_ParamAccess.item, 12);
            pManager.AddIntegerParameter("ALIGN", "", "", GH_ParamAccess.item, 1);

            pManager[1].Optional = true;
            Param_Integer param = (Param_Integer)pManager[4];
            param.AddNamedValue("Right", 0);
            param.AddNamedValue("Left", 1);
            param.AddNamedValue("Centre", 2);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("EXCEL MODS", "", "", GH_ParamAccess.list);
        }
        ExcelProperties excelProperties;
        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            List<string> cellNames = new List<string>();
            Color cellColour = Color.Azure;
            bool border = false;
            int fontSize = 12;
            int alignment = 2;

            if (!DA.GetDataList(0, cellNames)) return;
            DA.GetData(1, ref cellColour);
            if (!DA.GetData(2, ref border)) return;
            if (!DA.GetData(3, ref fontSize)) return;
            if (!DA.GetData(4, ref alignment)) return;

            excelProperties = new ExcelProperties();
            excelProperties.initialiseProperties(cellNames, cellColour, fontSize, border, alignment);

            DA.SetData(0, excelProperties);

        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon => GH_Excel_Tools.Properties.Resources.Icons_Settings;

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("967F9199-C086-4539-9B22-AB68366E7360"); }
        }
    }
}