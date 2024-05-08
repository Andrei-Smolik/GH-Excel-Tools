using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace GH_Excel_Tools
{
        internal class ExcelProperties
        {
            public List<string> cellNames = new List<string>();
            public Color cellColour = Color.Azure;
            public int fontSize = 12;
            public bool border = false;
            public int alignment = -1;

            public void initialiseProperties(List<string> cellNames, Color cellColour, int fontSize, bool border, int alignment)
            {
                this.cellNames = cellNames;
                this.cellColour = cellColour;
                this.fontSize = fontSize;
                this.border = border;
                this.alignment = alignment;
            }
        }
}
