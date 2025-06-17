using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Transformers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Objects.AutoCAD.Shapes
{
    public class Faces : AutoCADEntities
    {
        private Acad3DFace autoCADFace;

        public Faces(Acad3DFace face) : base((AcadEntity)face)
        {

            autoCADFace = face;

            // Initialize base class properties
            EntityType = face.EntityName;
            Handle = face.Handle;
            Layer = face.Layer;
            Color = face.TrueColor.ColorName;
            Linetype = face.Linetype;
            Lineweight = Convert.ToDouble(face.Lineweight);


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
            if (autoCADFace != null)
            {
                Marshal.ReleaseComObject(autoCADFace);
                autoCADFace = null;
            }
        }

    }
}
