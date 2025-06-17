using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Transformers;
using System.Runtime.InteropServices;


namespace CADInteropServices.Objects.AutoCAD.Annotations
{
    public class Hatches : AutoCADEntities
    {
        private AcadHatch autoCADHatch;

        public Hatches(AcadHatch hatch) : base((AcadEntity)hatch)
        {

            autoCADHatch = hatch;

            // Initialize base class properties
            EntityType = hatch.EntityName;
            Handle = hatch.Handle;
            Layer = hatch.Layer;
            Color = hatch.TrueColor.ColorName;
            Linetype = hatch.Linetype;
            Lineweight = Convert.ToDouble(hatch.Lineweight);


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
            if (autoCADHatch != null)
            {
                Marshal.ReleaseComObject(autoCADHatch);
                autoCADHatch = null;
            }
        }

    }
}
