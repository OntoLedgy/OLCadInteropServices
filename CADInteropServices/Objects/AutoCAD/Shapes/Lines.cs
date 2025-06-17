using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Objects.AutoCAD.Spaces;
using CADInteropServices.Transformers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Objects.AutoCAD.Shapes
{
    public class Lines : AutoCADEntities
    {
        private AcadLine acadLine;


        public Coordinates StartPoint { get; set; }
        public Coordinates EndPoint { get; set; }
        public double Angle { get; set; }

        public Lines(AcadLine lineEntity) : base((AcadEntity)lineEntity)
        {
            acadLine = lineEntity;

            // Initialize base class properties
            EntityType = acadLine.EntityName;
            Handle = acadLine.Handle;
            Layer = acadLine.Layer;
            Color = acadLine.TrueColor.ColorName;
            Linetype = acadLine.Linetype;
            Lineweight = Convert.ToDouble(acadLine.Lineweight);

            // Initialize specific properties
            StartPoint = new Coordinates(acadLine.StartPoint);
            EndPoint = new Coordinates(acadLine.EndPoint);
            Angle = acadLine.Angle;
        }

        public override string GetSpecificPropertiesAsString()
        {
            return $"StartPoint: {StartPoint}; EndPoint: {EndPoint}; Angle: {Angle}";
        }

        public override void Transform(TransformationMatrix matrix)
        {
            StartPoint = matrix.TransformCoordinates(StartPoint);
            EndPoint = matrix.TransformCoordinates(EndPoint);
        }


        public override void Report()
        {
            base.Report();
            Console.WriteLine($"  StartPoint: {StartPoint}");
            Console.WriteLine($"  EndPoint: {EndPoint}");
            Console.WriteLine($"  Angle: {Angle}");
        }

        public override void Release()
        {
            if (acadLine != null)
            {
                Marshal.ReleaseComObject(acadLine);
                acadLine = null;
            }
        }


    }


}

