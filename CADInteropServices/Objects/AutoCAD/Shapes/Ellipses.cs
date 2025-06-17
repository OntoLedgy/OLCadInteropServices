using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Transformers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Objects.AutoCAD.Shapes
{
    public class Ellipses : AutoCADEntities

    {
        private AcadEllipse autoCADEllipse;

        public Ellipses(AcadEllipse ellipse) : base((AcadEntity)ellipse)
        {
            autoCADEllipse = ellipse;

            // Initialize base class properties
            EntityType = ellipse.EntityName;
            Handle = ellipse.Handle;
            Layer = ellipse.Layer;
            Color = ellipse.TrueColor.ColorName;
            Linetype = ellipse.Linetype;
            Lineweight = Convert.ToDouble(ellipse.Lineweight);

        }

        public override string GetSpecificPropertiesAsString()
        {
            return $"TBD";
        }

        public override void Transform(TransformationMatrix matrix)
        {
            //TODO
        }

        public override void Report()
        {
            base.Report();
            Console.WriteLine($"  TBD");
        }

        public override void Release()
        {
            if (autoCADEllipse != null)
            {
                Marshal.ReleaseComObject(autoCADEllipse);
                autoCADEllipse = null;
            }
        }
    }
}
