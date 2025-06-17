namespace CADInteropServices.Objects.AutoCAD.Spaces
{
    public class Coordinates
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }

        // Constructors
        public Coordinates(double x, double y, double z)
        {
            X = x;
            Y = y;
            Z = z;
        }

        public Coordinates(double[] point)
        {
            if (point == null)
                throw new ArgumentNullException(nameof(point));

            if (point.Length >= 3)
            {
                X = point[0];
                Y = point[1];
                Z = point[2];
            }
            else if (point.Length == 2)
            {
                X = point[0];
                Y = point[1];
                Z = 0;
            }
            else
            {
                throw new ArgumentException("Point array must have at least 2 elements.", nameof(point));
            }
        }

        // New constructor to handle object input (e.g., from COM interop)
        public Coordinates(object pointObj)
        {
            if (pointObj == null)
                throw new ArgumentNullException(nameof(pointObj));

            //Console.WriteLine($"Creating Coordinates from object of type: {pointObj.GetType()}");

            switch (pointObj)
            {
                case Coordinates coords:
                    // Handle Coordinates input
                    X = coords.X;
                    Y = coords.Y;
                    Z = coords.Z;
                    break;

                case double[] doubleArray:
                    // Handle double array input
                    if (doubleArray.Length >= 3)
                    {
                        X = doubleArray[0];
                        Y = doubleArray[1];
                        Z = doubleArray[2];
                    }
                    else if (doubleArray.Length == 2)
                    {
                        X = doubleArray[0];
                        Y = doubleArray[1];
                        Z = 0;
                    }
                    else
                    {
                        throw new ArgumentException("Double array must have at least 2 elements.", nameof(pointObj));
                    }
                    break;

                case object[] objectArray:
                    // Handle object array input
                    if (objectArray.Length >= 3)
                    {
                        X = ConvertToDouble(objectArray[0]);
                        Y = ConvertToDouble(objectArray[1]);
                        Z = ConvertToDouble(objectArray[2]);
                    }
                    else if (objectArray.Length == 2)
                    {
                        X = ConvertToDouble(objectArray[0]);
                        Y = ConvertToDouble(objectArray[1]);
                        Z = 0;
                    }
                    else
                    {
                        throw new ArgumentException("Object array must have at least 2 elements.", nameof(pointObj));
                    }
                    break;

                case Array pointArray:
                    // Handle general array input
                    if (pointArray.Length >= 3)
                    {
                        X = ConvertToDouble(pointArray.GetValue(0));
                        Y = ConvertToDouble(pointArray.GetValue(1));
                        Z = ConvertToDouble(pointArray.GetValue(2));
                    }
                    else if (pointArray.Length == 2)
                    {
                        X = ConvertToDouble(pointArray.GetValue(0));
                        Y = ConvertToDouble(pointArray.GetValue(1));
                        Z = 0;
                    }
                    else
                    {
                        throw new ArgumentException("Point array must have at least 2 elements.", nameof(pointObj));
                    }
                    break;

                default:
                    throw new InvalidCastException($"Unable to convert object of type {pointObj.GetType()} to Coordinates.");
            }
        }


        // Helper method to safely convert objects to double
        private double ConvertToDouble(object value)
        {
            if (value == null)
                return 0.0;

            return Convert.ToDouble(value);
        }

        // Static method for conversion, if needed elsewhere
        public static Coordinates ConvertToCoordinates(object pointObj)
        {
            return new Coordinates(pointObj);
        }

        // Method to get coordinates as a double array
        public double[] GetCoordinates()
        {
            return new double[] { X, Y, Z };
        }

        public override string ToString()
        {
            return $"({X}, {Y}, {Z})";
        }
    }
}

