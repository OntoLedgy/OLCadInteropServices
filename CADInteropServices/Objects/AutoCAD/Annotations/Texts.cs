using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Objects.AutoCAD.Spaces;
using CADInteropServices.Transformers;
using System.Runtime.InteropServices;


namespace CADInteropServices.Objects.AutoCAD.Annotations
{
    public class Texts : AutoCADEntities
    {
        private AcadEntity textEntity;
        private AcadText acadTextEntity;
        private AcadMText acadMTextEntity;
        private AcadTextStyle acadTextStyleEntity;

        public string TextString { get; set; }
        public Coordinates InsertionPoint { get; set; }
        public double Height { get; set; }

        public Texts(AcadEntity textEntity) : base(textEntity)
        {
            acadEntity = textEntity;
            this.textEntity = textEntity;

            // Initialize base class properties
            EntityType = textEntity.EntityName;
            Handle = textEntity.Handle;
            Layer = textEntity.Layer;
            Color = textEntity.TrueColor.ColorName;
            Linetype = textEntity.Linetype;
            Lineweight = Convert.ToDouble(textEntity.Lineweight);
            BoundingBox = GetBoundingBox();

            switch (textEntity)
            {
                case AcadText acadTextEntity:
                    TextString = acadTextEntity.TextString;
                    InsertionPoint = new Coordinates(acadTextEntity.InsertionPoint);
                    Height = acadTextEntity.Height;
                    break;

                case AcadMText acadMTextEntity:
                    TextString = acadMTextEntity.TextString;
                    InsertionPoint = new Coordinates(acadMTextEntity.InsertionPoint);
                    Height = acadMTextEntity.Height;
                    break;

                case AcadTextStyle acadTextStyleEntity:
                    TextString = acadTextStyleEntity.Name;
                    break;



                // Add cases for additional entity types here
                default:
                    throw new InvalidOperationException("The provided entity is not a valid text.");
            }


        }

        public override string GetSpecificPropertiesAsString()
        {
            return $"TextString: {TextString}; InsertionPoint: {InsertionPoint}; Height: {Height}";
        }
        public override void Transform(TransformationMatrix matrix)
        {
            //TODO
        }

        public override void Report()
        {
            base.Report();
            Console.WriteLine($"  TextString: {TextString}");
            Console.WriteLine($"  InsertionPoint: {InsertionPoint}");
            Console.WriteLine($"  Height: {Height}");
        }

        public override void Release()
        {
            if (textEntity != null)
            {
                Marshal.ReleaseComObject(textEntity);
                textEntity = null;
            }
        }


    }
}
