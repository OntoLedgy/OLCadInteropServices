using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Objects.AutoCAD.Spaces;
using CADInteropServices.Transformers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Objects.AutoCAD.Shapes
{
    public class Circles : AutoCADEntities
    {
        private AcadCircle circleEntity;
        public Coordinates Centre { get; set; }
        public double Radius { get; set; }

        public Circles(AcadCircle circleEntity) : base((AcadEntity)circleEntity)
        {
            this.circleEntity = circleEntity;

            // Initialize base class properties
            EntityType = circleEntity.EntityName;
            Handle = circleEntity.Handle;
            Layer = circleEntity.Layer;
            Color = circleEntity.TrueColor.ColorName;
            Linetype = circleEntity.Linetype;
            Lineweight = Convert.ToDouble(circleEntity.Lineweight);

            // Initialize specific properties
            Centre = new Coordinates(circleEntity.Center);
            Radius = circleEntity.Radius;
        }

        public override string GetSpecificPropertiesAsString()
        {
            return $"Centre: {Centre}; Radius: {Radius}";
        }

        public override void Transform(TransformationMatrix matrix)
        {
            //TODO
        }

        public override void Report()
        {
            base.Report();
            Console.WriteLine($"  Centre: {Centre}");
            Console.WriteLine($"  Radius: {Radius}");
        }

        public override void Release()
        {
            if (circleEntity != null)
            {
                Marshal.ReleaseComObject(circleEntity);
                circleEntity = null;
            }
        }
    }

}
