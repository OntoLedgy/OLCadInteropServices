using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Transformers;
using System.Runtime.InteropServices;


namespace CADInteropServices.Objects.AutoCAD
{
    public class UnrecognizedEntities : AutoCADEntities
    {
        private AcadEntity? entity;

        public UnrecognizedEntities(AcadEntity entity) : base(entity)
        {
            this.entity = entity;

            // Initialize base class properties
            EntityType = entity.EntityName;
            Handle = entity.Handle;
            Layer = entity.Layer;
            Color = entity.TrueColor.ColorName;
            Linetype = entity.Linetype;
            Lineweight = Convert.ToDouble(entity.Lineweight);
        }

        public override void Report()
        {
            base.Report();
            Console.WriteLine($"  Unrecognized entity type: {EntityType}");
        }

        public override string GetSpecificPropertiesAsString()
        {
            return $"Unrecognized entity type: {EntityType}";
        }
        public override void Transform(TransformationMatrix matrix)
        {
            //TODO
        }
        public override void Release()
        {
            if (entity != null)
            {
                Marshal.ReleaseComObject(entity);
                entity = null;
            }
        }
    }

}
