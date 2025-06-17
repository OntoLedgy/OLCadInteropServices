using Autodesk.AutoCAD.Interop.Common;
using CADInteropServices.Objects.AutoCAD.Shapes;
using CADInteropServices.Objects.AutoCAD.Spaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static System.Reflection.Metadata.BlobBuilder;

namespace CADInteropServices.Transformers
{
    public class CoordinateTransformers
    {
        private void TransformLineEntity(
            Lines lineEntity, 
            TransformationMatrix transformationMatrix)
        {
            double[] transformedStart = transformationMatrix.TransformPoint(
                new double[] { lineEntity.StartPoint.X, lineEntity.StartPoint.Y, lineEntity.StartPoint.Z });

            double[] transformedEnd = transformationMatrix.TransformPoint(
                new double[] { lineEntity.EndPoint.X, lineEntity.EndPoint.Y, lineEntity.EndPoint.Z });

            lineEntity.StartPoint = new Coordinates(transformedStart[0], transformedStart[1], transformedStart[2]);
            lineEntity.EndPoint = new Coordinates(transformedEnd[0], transformedEnd[1], transformedEnd[2]);
        }
        private void TransformPolylineEntity(PolyLines polylineEntity, TransformationMatrix transformationMatrix)
        {
            for (int i = 0; i < polylineEntity.Vertices.Count; i++)
            {
                var vertex = polylineEntity.Vertices[i];
                double[] transformedVertex = transformationMatrix.TransformPoint(
                    new double[] { vertex.X, vertex.Y, vertex.Z });
                polylineEntity.Vertices[i] = new Coordinates(transformedVertex[0], transformedVertex[1], transformedVertex[2]);
            }
        }
        //private void TransformCircleEntity(Circles circleEntity, TransformationMatrix transformationMatrix)
        //{
        //    double[] transformedCenter = transformationMatrix.TransformPoint(
        //        new double[] { circleEntity.Center.X, circleEntity.Center.Y, circleEntity.Center.Z });
        //    circleEntity.Center = new Coordinates(transformedCenter[0], transformedCenter[1], transformedCenter[2]);
        //    // Handle scaling of radius if necessary
        //}

        //private void TransformBlockReferenceEntity(
        //        BlockReferences blockRef,
        //        TransformationMatrix parentTransformationMatrix)
        //{
        //    // Get the block definition by name
        //    AcadBlock blockDefinition = blocks.Item(blockRef.Name);

        //    // Create the transformation matrix for this block reference
        //    double[] insertionPoint = new double[]
        //    {
        //blockRef.InsertionPoint.X,
        //blockRef.InsertionPoint.Y,
        //blockRef.InsertionPoint.Z
        //    };
        //    double rotation = blockRef.Rotation;
        //    double[] scaleFactors = new double[]
        //    {
        //blockRef.XScaleFactor,
        //blockRef.YScaleFactor,
        //blockRef.ZScaleFactor
        //    };

        //    var transformationMatrix = new TransformationMatrix(
        //        insertionPoint,
        //        rotation,
        //        scaleFactors);

        //    // Combine with the parent transformation matrix
        //    var combinedTransformationMatrix = parentTransformationMatrix.Combine(transformationMatrix);

        //    // Iterate over entities in the block definition
        //    foreach (AcadEntity entity in blockDefinition)
        //    {
        //        // Recursively transform nested entities
        //        EntityBase transformedEntity = TransformEntity(entity, combinedTransformationMatrix);
        //        if (transformedEntity != null)
        //        {
        //            blockRef.TransformedEntities.Add(transformedEntity);
        //        }

        //        // Release the entity COM object
        //        Marshal.ReleaseComObject(entity);
        //    }

        //    // Release the block definition COM object
        //    Marshal.ReleaseComObject(blockDefinition);
        //}


    }
}
