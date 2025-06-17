using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects.AutoCAD.Annotations;
using CADInteropServices.Objects.AutoCAD.Spaces;
using CADInteropServices.Transformers;
using System.Runtime.InteropServices;


namespace CADInteropServices.Objects.AutoCAD
{

    public abstract class AutoCADEntities
    {
        //TODO replace this with EntityCommonInformation class
        protected AcadEntity acadEntity;

        public string EntityType { get; set; }

        public string Handle { get; set; }

        public string Layer { get; set; }

        public string Color { get; set; }

        public string Linetype { get; set; }

        public double Lineweight { get; set; }

        public BoundingBoxes BoundingBox { get; protected set; }

        // Base class constructor
        protected AutoCADEntities(
            AcadEntity acadEntity
            )
        {
            if (acadEntity == null)
            {
                throw new ArgumentNullException(nameof(acadEntity));
            }

            this.acadEntity = acadEntity;


            // Initialize general attributes
            EntityType = acadEntity.EntityName;
            Handle = acadEntity.Handle;
            Layer = acadEntity.Layer;
            Color = acadEntity.TrueColor?.ColorName ?? "ByLayer";
            Linetype = acadEntity.Linetype;
            Lineweight = Convert.ToDouble(acadEntity.Lineweight);

            // Initialize the bounding box
            InitializeBoundingBox();
        }

        public virtual void InitializeBoundingBox()
        {
            try
            {
                acadEntity.GetBoundingBox(out object minPointObj, out object maxPointObj);

                var minPoint = new Coordinates(minPointObj);
                var maxPoint = new Coordinates(maxPointObj);

                BoundingBox = new BoundingBoxes(minPoint, maxPoint);
            }
            catch (COMException comEx)
            {
                Console.WriteLine($"Entity '{EntityType}' does not have a valid bounding box. Skipping. COMException: {comEx.Message}");
                BoundingBox = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception in InitializeBoundingBox: {ex.Message}");
                BoundingBox = null;
            }
        }

        public virtual string GetSpecificPropertiesAsString()
        {
            return string.Empty;
        }

        public virtual BoundingBoxes GetBoundingBox()
        {
            if (acadEntity == null)
            {
                throw new InvalidOperationException("The underlying AcadEntity is null.");
            }

            try
            {
                acadEntity.GetBoundingBox(out object minPointObj, out object maxPointObj);

                double[] minPointArray = (double[])minPointObj;
                double[] maxPointArray = (double[])maxPointObj;

                var minPoint = new Coordinates(minPointArray[0], minPointArray[1], minPointArray[2]);
                var maxPoint = new Coordinates(maxPointArray[0], maxPointArray[1], maxPointArray[2]);

                return new BoundingBoxes(minPoint, maxPoint);
            }
            catch (COMException comEx)
            {
                Console.WriteLine($"COMException in GetBoundingBox: {comEx.Message}");
                return null;
            }
        }

        public abstract void Transform(TransformationMatrix matrix);

        public virtual void Report()
        {
            Console.WriteLine($"Entity Type: {EntityType}");
            Console.WriteLine($"  Handle: {Handle}");
            Console.WriteLine($"  Layer: {Layer}");
            Console.WriteLine($"  Color: {Color}");
            Console.WriteLine($"  Linetype: {Linetype}");
            Console.WriteLine($"  Lineweight: {Lineweight}");
        }

        public abstract void Release();

    }
}
