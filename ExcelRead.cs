using Grasshopper.Kernel.Data;
using Grasshopper.Kernel;
using Grasshopper;
using System.Collections.Generic;
using System.IO;
using System;
using ClosedXML.Excel;


namespace GH_Excel_Tools
{
    public class ExcelRead : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the ExcelRead class.
        /// </summary>
        public ExcelRead()
          : base("ExcelRead", "eRead",
              "Read Excel File",
              "X", "Data")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("FILE", "FILE", "", GH_ParamAccess.item);//file path
            pManager.AddTextParameter("SHEET", "SHEET", "", GH_ParamAccess.item);//sheet
            pManager.AddBooleanParameter("REFRESH", "", "", GH_ParamAccess.item);
            pManager[1].Optional = true;
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("DATA", "DATA", "", GH_ParamAccess.tree);//data
            pManager.AddTextParameter("SHEETS", "SHEETS", "", GH_ParamAccess.list);//data
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string filePath = "";
            string sheetName = "";
            DataTree<string> data = new DataTree<string>();
            bool refresh = false;

            if (!DA.GetData(0, ref filePath)) return;
            DA.GetData(1, ref sheetName);
            DA.GetData(2, ref refresh);

            XLWorkbook workbook = new XLWorkbook();
            //Catch for open Files:

            bool fileExists = File.Exists(filePath);
            Stream stream;

            if (!fileExists)
            {
                System.Windows.Forms.MessageBox.Show("File Does Not Exist!");
                return;
            }
            else
            {
                stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                workbook = new XLWorkbook(stream);

            }
            //Stream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            //workbook = new XLWorkbook(stream);
            var ws = workbook.Worksheet(1);

            //get worksheet if requested:
            if (worksheetExists(workbook, sheetName))
            {
                ws = workbook.Worksheet(sheetName);
            }
            else
            {
                if (sheetName != "") AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Worksheet Specified does not exist");
            }

            //get data:
            var range = ws.Range(ws.FirstCellUsed(), ws.LastCellUsed());
            var colCount = range.ColumnCount();
            var rowCount = range.RowCount();
            var colStart = ws.FirstColumnUsed().ColumnNumber();
            var rowStart = ws.FirstRowUsed().RowNumber();

            for (var c = colStart; c <= colCount + colStart - 1; c++)
            {
                for (var r = rowStart; r <= rowCount + rowStart - 1; r++)
                {
                    string value = ws.Cell(r, c).Value.ToString();
                    data.Add(value, new GH_Path(c));
                }
            }


            //debug
            List<string> debugWorksheetNames = new List<string>();
            for (int w = 1; w < workbook.Worksheets.Count + 1; w++)
            {
                //int fisrtCellCount = ws.FirstColumn().CellCount();
                debugWorksheetNames.Add(workbook.Worksheet(w).Name);

            }


            stream.Close();
            DA.SetDataTree(0, data);
            DA.SetDataList(1, debugWorksheetNames);//debug
        }

        bool worksheetExists(IXLWorkbook workBook, string sheetName)
        {
            bool result = false;
            for (int i = 1; i < workBook.Worksheets.Count + 1; i++)
            {
                //if (workBook.Worksheet(i) == null) break;
                if (workBook.Worksheet(i).Name == sheetName) result = true;
            }
            return result;
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                return GH_Excel_Tools.Properties.Resources.Icons_Read;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("607BA79A-85F9-4CE1-9EB9-7FD593E79C8B"); }
        }
    }
}