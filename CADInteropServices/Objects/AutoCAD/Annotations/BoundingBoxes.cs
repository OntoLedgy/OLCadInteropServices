using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CADInteropServices.Objects.AutoCAD.Spaces;

namespace CADInteropServices.Objects.AutoCAD.Annotations
{
    public class BoundingBoxes
    {
        public Coordinates MinPoint { get; set; }
        public Coordinates MaxPoint { get; set; }

        public BoundingBoxes(Coordinates minPoint, Coordinates maxPoint)
        {
            MinPoint = minPoint;
            MaxPoint = maxPoint;
        }

        public BoundingBoxes Combine(BoundingBoxes other)
        {
            double minX = Math.Min(MinPoint.X, other.MinPoint.X);
            double minY = Math.Min(MinPoint.Y, other.MinPoint.Y);
            double minZ = Math.Min(MinPoint.Z, other.MinPoint.Z);

            double maxX = Math.Max(MaxPoint.X, other.MaxPoint.X);
            double maxY = Math.Max(MaxPoint.Y, other.MaxPoint.Y);
            double maxZ = Math.Max(MaxPoint.Z, other.MaxPoint.Z);

            return new BoundingBoxes(
                new Coordinates(minX, minY, minZ),
                new Coordinates(maxX, maxY, maxZ)
            );
        }
    }
}
