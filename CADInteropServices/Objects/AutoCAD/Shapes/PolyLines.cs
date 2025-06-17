using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Objects.AutoCAD.Spaces;
using CADInteropServices.Transformers;

namespace CADInteropServices.Objects.AutoCAD.Shapes
{
    public class PolyLines : AutoCADEntities
    {
        private AcadEntity polylineEntity;
        private AcadPolyline acadPolyline;
        private AcadLWPolyline acadLWPolyline;

        public List<Coordinates> Vertices { get; set; }
        public bool Closed { get; set; }

        // Constructor for AcadLWPolyline
        public PolyLines(AcadLWPolyline lwPolyline) : base((AcadEntity)lwPolyline)
        {
            // Initialize base class properties
            EntityType = lwPolyline.EntityName;
            Handle = lwPolyline.Handle;
            Layer = lwPolyline.Layer;
            Color = lwPolyline.TrueColor.ColorName;
            Linetype = lwPolyline.Linetype;
            Lineweight = Convert.ToDouble(lwPolyline.Lineweight);

            // Initialize specific properties
            Vertices = new List<Coordinates>();
            Closed = lwPolyline.Closed;

            // Get the coordinates array
            object coordinatesObj = lwPolyline.Coordinates;

            if (coordinatesObj is Array coordinatesArray)
            {
                for (int i = 0; i < coordinatesArray.Length; i += 2)
                {
                    double x = Convert.ToDouble(coordinatesArray.GetValue(i));
                    double y = Convert.ToDouble(coordinatesArray.GetValue(i + 1));
                    Coordinates vertex = new Coordinates(x, y, 0);
                    Vertices.Add(vertex);
                }
            }
            else
            {
                Console.WriteLine("Unexpected coordinate array type for AcadLWPolyline.");
            }
        }

        // Constructor for AcadPolyline
        public PolyLines(AcadPolyline acadPolyline) : base((AcadEntity)acadPolyline)
        {
            // Initialize base class properties
            EntityType = acadPolyline.EntityName;
            Handle = acadPolyline.Handle;
            Layer = acadPolyline.Layer;
            Color = acadPolyline.TrueColor.ColorName;
            Linetype = acadPolyline.Linetype;
            Lineweight = Convert.ToDouble(acadPolyline.Lineweight);

            // Initialize specific properties
            Vertices = new List<Coordinates>();
            Closed = acadPolyline.Closed;

            // Get the coordinates array
            object coordinatesObj = acadPolyline.Coordinates;

            if (coordinatesObj is Array coordinatesArray)
            {
                for (int i = 0; i < coordinatesArray.Length; i += 3)
                {
                    double x = Convert.ToDouble(coordinatesArray.GetValue(i));
                    double y = Convert.ToDouble(coordinatesArray.GetValue(i + 1));
                    double z = Convert.ToDouble(coordinatesArray.GetValue(i + 2));
                    Coordinates vertex = new Coordinates(x, y, z);
                    Vertices.Add(vertex);
                }
            }
            else
            {
                Console.WriteLine("Unexpected coordinate array type for AcadPolyline.");
            }
        }

        public override string GetSpecificPropertiesAsString()
        {
            return $"Vertices: {string.Join("; ", Vertices)}; Closed: {Closed}";
        }

        public override void Transform(TransformationMatrix matrix)
        {
            //TODO
        }

        public override void Report()
        {
            base.Report();
            Console.WriteLine($"  Closed: {Closed}");
            Console.WriteLine("  Vertices:");
            foreach (var vertex in Vertices)
            {
                Console.WriteLine($"    {vertex}");
            }
        }

        public override void Release()
        {
            if (acadPolyline != null)
            {
                Marshal.ReleaseComObject(acadPolyline);
                acadPolyline = null;
            }

            if (acadLWPolyline != null)
            {
                Marshal.ReleaseComObject(acadLWPolyline);
                acadLWPolyline = null;
            }

            if (polylineEntity != null)
            {
                Marshal.ReleaseComObject(polylineEntity);
                polylineEntity = null;
            }
        }
    }
}
