using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD.Interop;
using CADInteropServices.Factories;
using CADInteropServices.Transformers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using CADInteropServices.Objects.AutoCAD;

namespace CADInteropServices.Objects.AutoCAD.Spaces
{
    public class Viewports : AutoCADEntities
    {
        private AcadPViewport viewportEntity;
        private AcadBlocks blocks;
        public Coordinates Center { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double TwistAngle { get; set; }
        public double CustomScale { get; set; }
        public Coordinates Target { get; set; }
        public List<AutoCADEntities> ModelEntities { get; private set; }

        public Viewports(AcadPViewport viewport, AcadBlocks blocks) : base((AcadEntity)viewport)
        {
            viewportEntity = viewport;
            this.blocks = blocks;

            // Initialize properties
            EntityType = viewport.EntityName;
            Handle = viewport.Handle;
            Layer = viewport.Layer;
            Color = viewport.TrueColor?.ColorName;
            Linetype = viewport.Linetype;
            Lineweight = Convert.ToDouble(viewport.Lineweight);

            Center = new Coordinates((double[])viewport.Center);
            Width = viewport.Width;
            Height = viewport.Height;
            TwistAngle = viewport.TwistAngle;
            CustomScale = viewport.CustomScale;
            Target = new Coordinates((double[])viewport.Target);

            ModelEntities = new List<AutoCADEntities>();

            if (viewport.ViewportOn)
            {
                GetModelSpaceEntities();
            }
        }

        private void GetModelSpaceEntities()
        {
            AcadDocument acadDoc = (AcadDocument)viewportEntity.Application.ActiveDocument;
            AcadModelSpace modelSpace = acadDoc.ModelSpace;

            TransformationMatrix transformationMatrix = CreateViewportTransformationMatrix();

            foreach (AcadEntity entity in modelSpace)
            {
                AutoCADEntities entityWrapper = EntityFactories.Create(entity, blocks);

                if (entityWrapper != null)
                {
                    // Apply the transformation to the entity
                    entityWrapper.Transform(transformationMatrix);
                    ModelEntities.Add(entityWrapper);
                }

                // Release the COM entity
                Marshal.ReleaseComObject(entity);
            }

            // Release the COM object
            Marshal.ReleaseComObject(modelSpace);
        }

        private TransformationMatrix CreateViewportTransformationMatrix()
        {
            // Use CustomScale as the scale factor
            double scale = CustomScale;

            // Get the target point (view center in model space)
            Coordinates viewCenter = Target;

            // Build the transformation matrix
            TransformationMatrix matrix = new TransformationMatrix();
            matrix.Translate(-viewCenter.X, -viewCenter.Y, -viewCenter.Z);
            matrix.RotateZ(-TwistAngle);
            matrix.Scale(scale, scale, scale);
            matrix.Translate(Center.X, Center.Y, Center.Z);

            return matrix;
        }

        public override void Transform(TransformationMatrix matrix)
        {
            Center = matrix.TransformCoordinates(Center);
        }

        public override void Release()
        {
            if (viewportEntity != null)
            {
                Marshal.ReleaseComObject(viewportEntity);
                viewportEntity = null;
            }
        }
    }

}
