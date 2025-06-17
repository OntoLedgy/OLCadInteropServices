using Autodesk.AutoCAD.Interop.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Objects.AutoCAD
{

    public class AutoCADDocumentLayers
    {
        AcadLayer _layer = null;
        public string Name { get; set; }
        public bool IsOn { get; set; }
        public bool IsFrozen { get; set; }
        public bool IsLocked { get; set; }
        public string ColourName { get; set; }
        public AcadAcCmColor TrueColour { get; set; } // RGB values


        public AutoCADDocumentLayers(
            string name,
            bool isOn,
            bool isFrozen,
            bool isLocked,
            string colourName,
            AcadAcCmColor trueColour)
        {
            Name = name;
            IsOn = isOn;
            IsFrozen = isFrozen;
            IsLocked = isLocked;
            ColourName = colourName;
            TrueColour = trueColour;
        }

        public void Release()
        {
            if (_layer != null)
            {
                Marshal.ReleaseComObject(_layer);
                _layer = null;
            }
        }


        public void Report()
        {
            Console.WriteLine($"Layer Name: {Name}");
            Console.WriteLine($"  Is On: {IsOn}");
            Console.WriteLine($"  Is Frozen: {IsFrozen}");
            Console.WriteLine($"  Is Locked: {IsLocked}");
            Console.WriteLine($"  Color Name: {ColourName}");
            Console.WriteLine($"  True Color: RGB({TrueColour.Red}, {TrueColour.Green}, {TrueColour.Blue})");
            Console.WriteLine();
        }

    }

}




