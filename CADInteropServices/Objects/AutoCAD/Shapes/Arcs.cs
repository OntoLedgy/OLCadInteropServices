using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Objects.AutoCAD.Spaces;
using CADInteropServices.Transformers;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Reflection.Emit;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Objects.AutoCAD.Shapes
{
    public class Arcs : AutoCADEntities
    {
        private AcadArc arc;
        public double Radius;
        public double ArcAngle;
        public Coordinates StartPoint;
        public Coordinates EndPoint;

        public Arcs(AcadArc arcentity) : base((AcadEntity)arcentity)
        {
            arc = arcentity;

            // Initialize base class properties
            EntityType = arcentity.EntityName;
            Handle = arcentity.Handle;
            Layer = arcentity.Layer;
            Color = arcentity.TrueColor.ColorName;
            Linetype = arcentity.Linetype;
            Lineweight = Convert.ToDouble(arcentity.Lineweight);

            Radius = Convert.ToDouble(arcentity.Radius);
            ArcAngle = arcentity.TotalAngle;
            StartPoint = new Coordinates(arcentity.StartPoint);
            EndPoint = new Coordinates(arcentity.EndPoint);

        }

        public override void Report()
        {
            base.Report();
            Console.WriteLine($"Radius: {Radius} ");
            Console.WriteLine($"ArcAngle: {ArcAngle} ");
            Console.WriteLine($"StartPoint: {StartPoint} ");
            Console.WriteLine($"EndPoint: {EndPoint} ");
        }

        public override string GetSpecificPropertiesAsString()
        {
            return $"Radius: {Radius}";
        }

        public override void Transform(TransformationMatrix matrix)
        {
            //TODO
        }

        public override void Release()
        {
            if (arc != null)
            {
                Marshal.ReleaseComObject(arc);
                arc = null;
            }
        }
    }
}

