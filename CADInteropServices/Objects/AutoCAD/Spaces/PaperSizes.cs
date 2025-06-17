using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Objects.AutoCAD.Spaces
{
    public class PaperSizes
    {
        public double Hieght { get; set; }
        public double Width { get; set; }

        public PaperSizes(
            double hieght,
            double width)
        {

            Hieght = hieght;
            Width = width;
        }

    }
}
