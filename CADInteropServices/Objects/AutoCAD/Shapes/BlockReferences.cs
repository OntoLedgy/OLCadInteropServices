using Autodesk.AutoCAD.Interop.Common;
using System.Runtime.InteropServices;
using CADInteropServices.Transformers;
using CADInteropServices.Factories;
using Newtonsoft.Json;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Objects.AutoCAD.Annotations;
using CADInteropServices.Objects.AutoCAD.Spaces;


namespace CADInteropServices.Objects.AutoCAD.Shapes
{
    public class BlockReferences : AutoCADEntities
    {

        public AcadBlocks blocks; // Reference to the Blocks collection
        //public AcadBlockReference acadEntity;
        public string DefinitionHandle;

        public string Name { get; set; }

        public Coordinates InsertionPoint { get; set; }


        public double Rotation { get; set; }


        public double XScaleFactor { get; set; }

        [JsonProperty]
        public double YScaleFactor { get; set; }

        [JsonProperty]
        public double ZScaleFactor { get; set; }

        [JsonIgnore]
        public List<AutoCADEntities> TransformedEntities { get; set; }
        [JsonIgnore]
        private TransformationMatrix transformationMatrix;


        public BlockReferences(
            AcadBlockReference blockReferenceEntity,
            AcadBlocks blocks,
            TransformationMatrix parentTransformationMatrix = null)
            : base((AcadEntity)blockReferenceEntity)
        {

            this.blocks = blocks;

            // Initialize base class properties
            EntityType = blockReferenceEntity.EntityName;
            Handle = blockReferenceEntity.Handle;
            Layer = blockReferenceEntity.Layer;
            Color = blockReferenceEntity.TrueColor.ColorName;
            Linetype = blockReferenceEntity.Linetype;
            Lineweight = Convert.ToDouble(blockReferenceEntity.Lineweight);


            // Initialize specific properties
            Name = GetEffectiveBlockName(blockReferenceEntity);


            InsertionPoint = new Coordinates(blockReferenceEntity.InsertionPoint);
            Rotation = blockReferenceEntity.Rotation;

            XScaleFactor = blockReferenceEntity.XScaleFactor;
            YScaleFactor = blockReferenceEntity.YScaleFactor;
            ZScaleFactor = blockReferenceEntity.ZScaleFactor;

            //TODO: extract the bounding box as an object            
            //blockReferenceEntity.GetBoundingBox(out double[] lowerLeft, out double[] upperRight);

            TransformedEntities = new List<AutoCADEntities>();

            //prepare transformation matrix
            // Create the local transformation matrix
            double[] insertionPoint =
            {
                InsertionPoint.X,
                InsertionPoint.Y,
                InsertionPoint.Z
            };
            double rotation = Rotation;

            double[] scaleFactors =
            {
                XScaleFactor,
                YScaleFactor,
                ZScaleFactor
            };

            var localTransformationMatrix = new TransformationMatrix(
                insertionPoint,
                rotation,
                scaleFactors);

            // Combine with parent transformation matrix if provided
            if (parentTransformationMatrix != null)
            {
                transformationMatrix = parentTransformationMatrix.Combine(localTransformationMatrix);
            }
            else
            {
                transformationMatrix = localTransformationMatrix;
            }

            // Get the block definition by name
            AcadBlock blockDefinition = blocks.Item(Name);

            // Store the block definition's Handle
            DefinitionHandle = blockDefinition.Handle;

            // Process the block definition
            ProcessBlockDefinition(blockDefinition);

            // Release the block definition COM object
            Marshal.ReleaseComObject(blockDefinition);

            InitializeBoundingBox();
        }

        private string GetEffectiveBlockName(AcadBlockReference blockRef)
        {
            try
            {
                if (blockRef.IsDynamicBlock)
                {
                    // For dynamic blocks, get the effective name
                    return blockRef.EffectiveName;
                }
                else
                {
                    return blockRef.Name;
                }
            }
            catch
            {
                return blockRef.Name;
            }
        }


        private void ProcessBlockDefinition(AcadBlock blockDefinition)
        {

            // Iterate over entities in the block definition
            foreach (AcadEntity entity in blockDefinition)
            {
                // Transform the entity's geometry
                AutoCADEntities transformedEntity = TransformEntity(
                    entity,
                    transformationMatrix);

                if (transformedEntity != null)
                {
                    TransformedEntities.Add(transformedEntity);
                }

                // Release the entity COM object
                Marshal.ReleaseComObject(entity);
            }

            // Release the block definition COM object
            Marshal.ReleaseComObject(blockDefinition);
        }


        private AutoCADEntities TransformEntity(
            AcadEntity entity,
            TransformationMatrix transformationMatrix)
        {
            // Create a wrapper for the entity
            AutoCADEntities entityWrapper = EntityFactories.Create(
                entity,
                blocks,
                transformationMatrix);

            if (entityWrapper == null)
            {
                return null;
            }

            if (entityWrapper is BlockReferences)
            {
                // The nested BlockReferences instance has already processed its block definition with the correct transformation matrix
                // No further action is needed
            }
            else
            {
                // Apply transformation to the entity
                TransformEntity(
                    entityWrapper,
                    transformationMatrix);
            }

            return entityWrapper;
        }

        private void TransformEntity(
            AutoCADEntities entityWrapper,
            TransformationMatrix transformationMatrix)
        {
            switch (entityWrapper)
            {
                case Lines lineEntity:
                    TransformLineEntity(
                        lineEntity,
                        transformationMatrix);
                    break;

                case PolyLines polylineEntity:
                    TransformPolylineEntity(
                        polylineEntity,
                        transformationMatrix);
                    break;

                case Arcs arc:
                    //TOOD: add transformation
                    break;

                case AttributeDefinitions attributeDefinitions:
                    //TODO add transformation
                    break;
                case Texts text:
                    //TODO add transformation
                    break;

                case Circles circle:
                    TransformCircleEntity(circle, transformationMatrix);
                    break;

                case Hatches hatch:
                    break;

                case Solids solid:
                    break;
                // Handle other entity types as needed

                case Ellipses ellipse:
                    break;

                case Faces face:
                    break;

                case Leaders leader:
                    break;
                case UnrecognizedEntities unrecognizedEntity:
                    // Optionally apply a generic transformation or leave as is
                    break;

                default:
                    Console.WriteLine($"Unhandled entity type: {entityWrapper.EntityType}");
                    break;
            }
        }


        public void ApplyTransformation(
            TransformationMatrix parentTransformationMatrix)
        {
            // Combine the current transformation matrix with the parent
            var combinedTransformationMatrix = parentTransformationMatrix.Combine(transformationMatrix);

            foreach (var entity in TransformedEntities)
            {
                if (entity is BlockReferences nestedBlockRef)
                {
                    // Recursively apply transformations
                    nestedBlockRef.ApplyTransformation(combinedTransformationMatrix);
                }
                else
                {
                    // Apply transformation to the entity
                    TransformEntity(entity, combinedTransformationMatrix);
                }
            }
        }

        // Transformation methods for different entity types
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

        private void TransformPolylineEntity(
            PolyLines polylineEntity,
            TransformationMatrix transformationMatrix)
        {
            for (int i = 0; i < polylineEntity.Vertices.Count; i++)
            {
                var vertex = polylineEntity.Vertices[i];
                double[] transformedVertex = transformationMatrix.TransformPoint(
                    new double[] { vertex.X, vertex.Y, vertex.Z });
                polylineEntity.Vertices[i] = new Coordinates(transformedVertex[0], transformedVertex[1], transformedVertex[2]);
            }
        }

        private void TransformCircleEntity(
            Circles circleEntity,
            TransformationMatrix transformationMatrix)
        {
            double[] transformedCenter = transformationMatrix.TransformPoint(
                new double[] { circleEntity.Centre.X, circleEntity.Centre.Y, circleEntity.Centre.Z });
            circleEntity.Centre = new Coordinates(transformedCenter[0], transformedCenter[1], transformedCenter[2]);
            // Handle scaling of radius if necessary
        }

        private void InitializeBoundingBox()
        {
            try
            {
                object minPointObj, maxPointObj;
                var blockRef = (AcadBlockReference)acadEntity;

                // Get the bounding box directly from the block reference
                blockRef.GetBoundingBox(out minPointObj, out maxPointObj);

                if (minPointObj == null || maxPointObj == null)
                {
                    Console.WriteLine($"Block reference {Handle} has null extents.");
                    BoundingBox = null;
                    return;
                }

                Coordinates minPoint = new Coordinates(minPointObj);
                Coordinates maxPoint = new Coordinates(maxPointObj);

                BoundingBox = new BoundingBoxes(minPoint, maxPoint);
            }
            catch (COMException comEx) when (comEx.Message.Contains("Null extents"))
            {
                Console.WriteLine($"Block reference {Handle} has null extents: {comEx.Message}");
                BoundingBox = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing bounding box for block reference {Handle}: {ex.Message}");
                BoundingBox = null;
            }
        }

        public override void Transform(TransformationMatrix matrix)
        {
            //TODO
        }

        public override void Report()
        {
            base.Report();
            Console.WriteLine($"  Name: {Name}");
            Console.WriteLine($"  InsertionPoint: {InsertionPoint}");
            Console.WriteLine($"  Rotation: {Rotation}");
        }

        public override string GetSpecificPropertiesAsString()
        {
            return $"Name: {Name}; InsertionPoint: {InsertionPoint}; Rotation: {Rotation}";
        }

        public override void Release()
        {
            if (acadEntity != null)
            {
                Marshal.ReleaseComObject(acadEntity);
                acadEntity = null;
            }

            // Release transformed entities
            if (TransformedEntities != null)
            {
                foreach (var entity in TransformedEntities)
                {
                    entity.Release();
                }
                TransformedEntities.Clear();
                TransformedEntities = null;
            }

            if (blocks != null)
            {
                Marshal.ReleaseComObject(blocks);
                blocks = null;
            }
        }
    }
}
