using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Transformers;
using System.Runtime.InteropServices;


namespace CADInteropServices.Objects.AutoCAD.Annotations
{
    public class Leaders : AutoCADEntities
    {
        private AcadLeader autoCADLeader;

        public Leaders(AcadLeader leader) : base((AcadEntity)leader)
        {

            autoCADLeader = leader;

            // Initialize base class properties
            EntityType = leader.EntityName;
            Handle = leader.Handle;
            Layer = leader.Layer;
            Color = leader.TrueColor.ColorName;
            Linetype = leader.Linetype;
            Lineweight = Convert.ToDouble(leader.Lineweight);


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
            if (autoCADLeader != null)
            {
                Marshal.ReleaseComObject(autoCADLeader);
                autoCADLeader = null;
            }
        }

    }
}
