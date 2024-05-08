using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace GH_Excel_Tools
{
    public class GH_Excel_ToolsInfo : GH_AssemblyInfo
    {
        public override string Name => "GH Excel Tools";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => null;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "";

        public override Guid Id => new Guid("6a63c803-9997-4469-a7d4-1e5c58d6d098");

        //Return a string identifying you or your company.
        public override string AuthorName => "";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "";
    }
}