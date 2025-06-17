using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Transformers;
using System.Runtime.InteropServices;

namespace CADInteropServices.Objects.AutoCAD
{
    public class AttributeDefinitions : AutoCADEntities
    {
        private AcadEntity attributeEntity;
        private AcadAttribute autoCADAttribute;
        private AcadAttributeReference AutoCADAttributereference;

        public AttributeDefinitions(AcadEntity attributeDefinition) : base(attributeDefinition)
        {
            attributeEntity = attributeDefinition;

            // Initialize base class properties
            EntityType = attributeDefinition.EntityName;
            Handle = attributeDefinition.Handle;
            Layer = attributeDefinition.Layer;
            Color = attributeDefinition.TrueColor.ColorName;
            Linetype = attributeDefinition.Linetype;
            Lineweight = Convert.ToDouble(attributeDefinition.Lineweight);


            switch (attributeEntity)
            {
                case AcadAttribute autoCADAttribute:
                    break;
                case AcadAttributeReference AutoCADAttributereference:
                    break;
                default:
                    Console.WriteLine($"Unrecognized attribute entity type: {EntityType}");
                    break;

            }


        }

        public override void InitializeBoundingBox()
        {

        }

        public override void Report()
        {
            base.Report();
            // to be added
        }
        public override void Transform(TransformationMatrix matrix)
        {
            //TODO
        }
        public override string GetSpecificPropertiesAsString()
        {
            return $"Unrecognized attribute entity type: {EntityType}";
        }

        public override void Release()
        {
            if (attributeEntity != null)
            {
                Marshal.ReleaseComObject(attributeEntity);
                attributeEntity = null;
            }
        }
    }
}

