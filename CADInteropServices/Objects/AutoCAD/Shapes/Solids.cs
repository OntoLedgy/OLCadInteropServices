using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Transformers;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Objects.AutoCAD.Shapes
{
    public class Solids : AutoCADEntities
    {
        private AcadSolid solidEntity;

        public Solids(AcadSolid solidEntity) : base((AcadEntity)solidEntity)
        {
            this.solidEntity = solidEntity;

            // Initialize base class properties
            EntityType = solidEntity.EntityName;
            Handle = solidEntity.Handle;
            Layer = solidEntity.Layer;
            Color = solidEntity.TrueColor.ColorName;
            Linetype = solidEntity.Linetype;
            Lineweight = Convert.ToDouble(solidEntity.Lineweight);

            // Initialize specific properties as needed
            // ...
        }

        public override void Transform(TransformationMatrix matrix)
        {
            //TODO
        }

        public override void Report()
        {
            base.Report();
            // Report specific properties
        }

        public override void Release()
        {
            if (solidEntity != null)
            {
                Marshal.ReleaseComObject(solidEntity);
                solidEntity = null;
            }
        }
    }


}
