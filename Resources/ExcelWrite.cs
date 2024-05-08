using System;
using System.Collections.Generic;
using System.IO;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using ClosedXML.Excel;//excel library

using Grasshopper;

namespace GH_Excel_Tools
{
    public class SaveExcel : GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public SaveExcel()
          : base("Save Excel", "eSave",
            "Save Excel File",
            "X", "Data")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("FILE", "FILE", "", GH_ParamAccess.item);
            pManager.AddGenericParameter("DATA", "DATA", "", GH_ParamAccess.tree);
            
            pManager.AddGenericParameter("MODIFY", "", "", GH_ParamAccess.list);
            pManager.AddTextParameter("SHEET", "", "", GH_ParamAccess.item, "Sheet1");
            pManager.AddBooleanParameter("SAVE", "", "", GH_ParamAccess.item);

            pManager[3].Optional = true;

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            //pManager.AddTextParameter("debug", "debug", "debug", GH_ParamAccess.item);//debugItem
            //pManager.AddTextParameter("debugList", "debugList", "debug2", GH_ParamAccess.list);//debugList
        }


        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GH_Structure<IGH_Goo> dataGoo;
            DataTree<string> data = new DataTree<string>();
            //List<string> debugStrings = new List<string>();//debug
            string fileName = "Excel File";
            bool saveState = false;
            List<ExcelProperties> excelPropertiesList = new List<ExcelProperties>();
            string sheetName = "Sheet1";

            if (!DA.GetData(0, ref fileName)) return;//file name
            if (!DA.GetDataTree(1, out dataGoo)) return;//data
            DA.GetDataList(2, excelPropertiesList);//properties

            DA.GetData(3, ref sheetName);//sheet name
            if (!DA.GetData(4, ref saveState)) return;

            //create Excel file:
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add(sheetName);

            for (int b = 0; b < dataGoo.Branches.Count; b++)
            {
                for (int i = 0; i < dataGoo.Branches[b].Count; i++)
                {
                    if (dataGoo.Branches[b][i] == null)
                    {
                        ws.Cell(i + 1, b + 1).Value = "";
                        continue;
                    }
                    //debugStrings.Add(dataGoo.Branches[b][i].TypeName);//debug
                    if (dataGoo.Branches[b][i].TypeName == "Text")
                        ws.Cell(i + 1, b + 1).Value = dataGoo.Branches[b][i].ToString();

                    if (dataGoo.Branches[b][i].TypeName == "Number")
                    {
                        double number = 0;
                        dataGoo.Branches[b][i].CastTo<double>(out number);
                        ws.Cell(i + 1, b + 1).Value = number;

                    }
                }
            }
            //custom modifiers:
            foreach (ExcelProperties excelProperty in excelPropertiesList)
                modifyWorkbook(ref ws, excelProperty);


            //Catch for open Files:
            if (saveState)
            {
                bool fileExists = File.Exists(fileName);

                if (!fileExists)
                {
                    workbook.SaveAs($@"{fileName}");
                    //goto FINISH;
                }
                else
                {
                    try
                    {
                        Stream stream = File.Open($@"{fileName}", FileMode.Open, FileAccess.Read, FileShare.None);
                        stream.Close();
                        workbook.SaveAs($@"{fileName}");
                    }
                    catch
                    {
                        System.Windows.Forms.MessageBox.Show("File Opened!");
                    }
                }
            }

            //DA.SetData(0, "Saved");
        }

        void modifyWorkbook(ref IXLWorksheet ws, ExcelProperties excelProperties)
        {
            foreach (string name in excelProperties.cellNames)
            {
                //XLCellValue xlCellVal=new XLCellValue();
                ClosedXML.Excel.IXLCell cell = ws.Cell(name);
                if (cell == null) continue;

                //cell background colour
                cell.Style.Fill.BackgroundColor = XLColor.FromColor(excelProperties.cellColour);

                //cell font size
                cell.Style.Font.FontSize = excelProperties.fontSize;

                //cell border
                if (excelProperties.border)
                {
                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                    cell.Style.Border.RightBorder = XLBorderStyleValues.Thick;
                    cell.Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                    cell.Style.Border.TopBorder = XLBorderStyleValues.Thick;
                }

                //cell alignment
                switch (excelProperties.alignment)
                {
                    case 0:
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        break;
                    case 1:
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        break;
                    case 2:
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        break;
                    default:
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        break;
                }
            }
        }

        DataTree<string> convertTree(GH_Structure<IGH_Goo> goo)
        {
            DataTree<string> dataTree = new DataTree<string>();

            for (int b = 0; b < goo.Branches.Count; b++)
            {
                for (int i = 0; i < goo.Branches[b].Count; i++)
                {
                    dataTree.Add(goo.Branches[b][i].ToString(), new GH_Path(b));
                }
            }

            return dataTree;
        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// You can add image files to your project resources and access them like this:
        /// return Resources.IconForThisComponent;
        /// </summary>
        protected override System.Drawing.Bitmap Icon => GH_Excel_Tools.Properties.Resources.Icons_Write;

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("27f4835f-347f-4a35-b8f2-2a4209ce9885");
    }
}