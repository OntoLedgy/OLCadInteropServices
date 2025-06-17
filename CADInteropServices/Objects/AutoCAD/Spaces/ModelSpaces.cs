using Autodesk.AutoCAD.Interop.Common;
using System.Runtime.InteropServices;

using CADInteropServices.Factories;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Objects.AutoCAD.Shapes;

namespace CADInteropServices.Objects.AutoCAD.Spaces
{
    public class ModelSpaces : IDisposable
    {
        private AcadModelSpace autoCADModelSpace;
        private AcadBlocks blocks;
        private List<AutoCADEntities>? entities;
        private List<AutoCADDocumentLayers>? layers;

        public ModelSpaces(
            AcadModelSpace modelSpace,
            AcadBlocks blocks)
        {
            autoCADModelSpace = modelSpace;
            entities = null;
            layers = null;
            this.blocks = blocks;
        }

        public List<AutoCADEntities> GetAllEntities()
        {
            if (entities != null)
            {
                return entities;
            }

            entities = new List<AutoCADEntities>();
            const int maxRetries = 10;
            int retryCount = 0;
            bool success = false;

            while (!success && retryCount < maxRetries)
            {
                try
                {
                    // Iterate over entities in ModelSpace
                    foreach (AcadEntity entity in autoCADModelSpace)
                    {
                        Console.WriteLine($"loaded entity: {entity.EntityName}, {entity.Handle}, {entity.ObjectID}");
                        AutoCADEntities entityWrapper = EntityFactories.Create(entity, blocks);

                        if (entityWrapper != null)
                        {

                            if (entityWrapper is BlockReferences blockReference)
                            {
                                Console.WriteLine($"loaded block entity: {blockReference.Name}");
                            }
                            entities.Add(entityWrapper);

                        }
                        else
                        {
                            // If we don't have a wrapper for this entity, release it
                            Console.WriteLine($"Cannot load entity because handle is not defined: {entity.EntityName}, {entity.Handle}");
                            _ = Marshal.ReleaseComObject(entity);
                        }
                    }

                    success = true;
                }
                catch (COMException comEx) when ((uint)comEx.ErrorCode == 0x8001010A)
                {
                    // RPC_E_SERVERCALL_RETRYLATER
                    retryCount++;
                    Console.WriteLine($"COM Exception: {comEx.Message}. Retrying {retryCount}/{maxRetries}...");
                    Thread.Sleep(500);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception: {ex.Message} StackTrack:{ex.StackTrace}");
                    break;
                }
            }

            if (!success)
            {
                Console.WriteLine("Failed to retrieve entities after multiple retries.");
            }

            return entities;
        }

        public List<AutoCADDocumentLayers> GetAllLayers()
        {
            if (layers != null)
            {
                return layers;
            }

            layers = new List<AutoCADDocumentLayers>();
            const int maxRetries = 10;
            int retryCount = 0;
            bool success = false;

            AcadLayers layersList = null;

            while (!success && retryCount < maxRetries)
            {
                try
                {
                    layersList = autoCADModelSpace.Application.ActiveDocument.Layers;

                    foreach (AcadLayer layer in layersList)
                    {
                        AutoCADDocumentLayers layerInfo = new AutoCADDocumentLayers(
                            name: layer.Name,
                            isOn: layer.LayerOn,
                            isFrozen: layer.Freeze,
                            isLocked: layer.Lock,
                            colourName: layer.TrueColor.ColorName,
                            trueColour: layer.TrueColor
                        );

                        layers.Add(layerInfo);

                        _ = Marshal.ReleaseComObject(layer);
                    }

                    _ = Marshal.ReleaseComObject(layersList);

                    success = true;
                }
                catch (COMException comEx) when ((uint)comEx.ErrorCode == 0x8001010A)
                {
                    // RPC_E_SERVERCALL_RETRYLATER
                    retryCount++;
                    Console.WriteLine($"COM Exception: {comEx.Message}. Retrying {retryCount}/{maxRetries}...");
                    Thread.Sleep(500);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception: {ex.Message}");
                    break; // Break on non-COM exceptions
                }
                finally
                {
                    if (layersList != null)
                    {
                        _ = Marshal.ReleaseComObject(layersList);
                        layersList = null;
                    }
                }
            }

            if (!success)
            {
                Console.WriteLine("Failed to retrieve layers after multiple retries.");
            }

            return layers;
        }

        public void ReportAllLayers()
        {
            List<AutoCADDocumentLayers> layers = GetAllLayers();

            Console.WriteLine("Layers in the ModelSpace:");
            foreach (var layer in layers)
            {
                layer.Report();
            }
        }

        public Dictionary<string, int> GetAllEntityTypes()
        {
            Dictionary<string, int> entityTypes = new Dictionary<string, int>();
            const int maxRetries = 10;
            int retryCount = 0;
            bool success = false;

            while (!success && retryCount < maxRetries)
            {
                try
                {
                    // Iterate over entities in ModelSpace
                    foreach (AcadEntity entity in autoCADModelSpace)
                    {
                        string entityType = entity.EntityName;

                        if (entityTypes.ContainsKey(entityType))
                        {
                            entityTypes[entityType]++;
                        }
                        else
                        {
                            entityTypes[entityType] = 1;
                        }

                        // Release COM object for the entity
                        _ = Marshal.ReleaseComObject(entity);
                    }

                    success = true;
                }
                catch (COMException comEx) when ((uint)comEx.ErrorCode == 0x8001010A)
                {
                    // RPC_E_SERVERCALL_RETRYLATER
                    retryCount++;
                    Console.WriteLine($"COM Exception: {comEx.Message}. Retrying {retryCount}/{maxRetries}...");
                    Thread.Sleep(500);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception: {ex.Message}");
                    break;
                }
            }

            if (!success)
            {
                Console.WriteLine("Failed to retrieve entity types after multiple retries.");
            }

            return entityTypes;
        }

        public List<AutoCADEntities> GetEntitiesByLayer(
            string layerName)
        {
            List<AutoCADEntities> allEntities = GetAllEntities();
            List<AutoCADEntities> entitiesOnLayer = new List<AutoCADEntities>();

            foreach (var entity in allEntities)
            {
                if (entity.Layer.Equals(layerName, StringComparison.OrdinalIgnoreCase))
                {
                    entitiesOnLayer.Add(entity);
                }
                else
                {
                    // Release entities not on the desired layer
                    entity.Release();
                }
            }

            return entitiesOnLayer;
        }

        public void ReportAllEntities()
        {
            List<AutoCADEntities> entities = GetAllEntities();

            foreach (var entity in entities)
            {
                entity.Report();

                // Release the entity COM object
                entity.Release();

                Console.WriteLine();
            }
        }

        public void ReportEntitiesByLayer(
            string layerName)
        {
            List<AutoCADEntities> entities = GetEntitiesByLayer(layerName);

            foreach (var entity in entities)
            {
                entity.Report();

                // Release the entity COM object
                entity.Release();

                Console.WriteLine();
            }
        }

        public void ReportAllEntityTypes()
        {
            var entityTypes = GetAllEntityTypes();

            Console.WriteLine("Entity Types in the ModelSpace:");
            foreach (var kvp in entityTypes)
            {
                Console.WriteLine($"Entity Type: {kvp.Key}, Count: {kvp.Value}");
            }
        }

        public void AddLayer(string layerName)
        {
            const int maxRetries = 10;
            int retryCount = 0;
            bool success = false;

            AcadLayers layersList = null;
            AcadLayer newLayer = null;

            while (!success && retryCount < maxRetries)
            {
                try
                {
                    layersList = autoCADModelSpace.Application.ActiveDocument.Layers;

                    // Check if the layer already exists
                    bool layerExists = false;
                    foreach (AcadLayer layer in layersList)
                    {
                        if (layer.Name.Equals(layerName, StringComparison.OrdinalIgnoreCase))
                        {
                            layerExists = true;
                            Console.WriteLine($"Layer '{layerName}' already exists.");
                            Marshal.ReleaseComObject(layer);
                            break;
                        }
                        Marshal.ReleaseComObject(layer);
                    }

                    if (!layerExists)
                    {
                        // Add the new layer
                        newLayer = layersList.Add(layerName);
                        Console.WriteLine($"Layer '{layerName}' added successfully.");
                    }

                    success = true;
                }
                catch (COMException comEx) when ((uint)comEx.ErrorCode == 0x8001010A)
                {
                    // RPC_E_SERVERCALL_RETRYLATER
                    retryCount++;
                    Console.WriteLine($"COM Exception: {comEx.Message}. Retrying {retryCount}/{maxRetries}...");
                    Thread.Sleep(500);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception: {ex.Message}");
                    break; // Break on non-COM exceptions
                }
                finally
                {
                    // Release COM objects
                    if (newLayer != null)
                    {
                        Marshal.ReleaseComObject(newLayer);
                        newLayer = null;
                    }

                    if (layersList != null)
                    {
                        Marshal.ReleaseComObject(layersList);
                        layersList = null;
                    }
                }
            }

            if (!success)
            {
                Console.WriteLine($"Failed to add layer '{layerName}' after multiple retries.");
            }
        }


        public void AddRectangle(
            Coordinates lowerLeft,
            Coordinates upperRight,
            string layerName = null)
        {
            const int maxRetries = 10;
            int retryCount = 0;
            bool success = false;

            AcadLWPolyline rectangle = null;

            // Validate coordinates
            if (lowerLeft == null || upperRight == null)
            {
                Console.WriteLine("Invalid coordinates provided to AddRectangle.");
                return;
            }

            // Ensure that the rectangle has positive dimensions
            if (lowerLeft.X == upperRight.X || lowerLeft.Y == upperRight.Y)
            {
                Console.WriteLine("Rectangle has zero width or height. Skipping.");
                return;
            }

            while (!success && retryCount < maxRetries)
            {
                try
                {
                    // Define the rectangle corners
                    double x1 = lowerLeft.X;
                    double y1 = lowerLeft.Y;

                    double x2 = upperRight.X;
                    double y2 = upperRight.Y;

                    // Define the points of the rectangle (as doubles)
                    double[] points = new double[]
                    {
                x1, y1,   // Lower Left
                x2, y1,   // Lower Right
                x2, y2,   // Upper Right
                x1, y2,   // Upper Left
                x1, y1    // Close the rectangle by returning to the Lower Left
                    };

                    // Add the lightweight polyline (rectangle) to the model space
                    rectangle = autoCADModelSpace.AddLightWeightPolyline(points);

                    if (rectangle != null)
                    {
                        // Close the polyline to make it a closed rectangle
                        rectangle.Closed = true;

                        // Set the layer if provided
                        if (!string.IsNullOrEmpty(layerName))
                        {
                            rectangle.Layer = layerName;
                        }

                        Console.WriteLine("Rectangle added successfully.");
                        autoCADModelSpace.Application.ActiveDocument.Regen(AcRegenType.acActiveViewport);
                    }
                    else
                    {
                        Console.WriteLine("Failed to add rectangle.");
                    }

                    success = true;
                }
                catch (COMException comEx) when ((uint)comEx.ErrorCode == 0x8001010A)
                {
                    // RPC_E_SERVERCALL_RETRYLATER
                    retryCount++;
                    Console.WriteLine($"COM Exception: {comEx.Message}. Retrying {retryCount}/{maxRetries}...");
                    Thread.Sleep(500);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception in AddRectangle: {ex.Message}");
                    break;
                }
                finally
                {
                    // Release the rectangle COM object
                    if (rectangle != null)
                    {
                        _ = Marshal.ReleaseComObject(rectangle);
                        rectangle = null;
                    }
                }
            }

            if (!success)
            {
                Console.WriteLine("Failed to add rectangle after multiple retries.");
            }
        }


        public void AddCircle(
            Coordinates center,
            double radius)
        {
            try
            {
                AcadCircle circle = autoCADModelSpace.AddCircle(
                    center.GetCoordinates(),
                    radius);

                if (circle != null)
                {
                    Console.WriteLine("Circle added successfully.");
                    autoCADModelSpace.Application.ActiveDocument.Regen(AcRegenType.acActiveViewport);
                }
                else
                {
                    Console.WriteLine("Failed to add circle.");
                }

                _ = Marshal.ReleaseComObject(circle);
            }
            catch (COMException comEx)
            {
                Console.WriteLine("COM Exception: " + comEx.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }
        }
        public void AddBoundingBoxRectangles(string newLayerName)
        {
            // Add the new layer
            AddLayer(newLayerName);

            // Get all entities in the model space
            List<AutoCADEntities> entities = GetAllEntities();

            foreach (var entity in entities)
            {
                if (entity.BoundingBox != null)
                {
                    Coordinates minPoint = entity.BoundingBox.MinPoint;
                    Coordinates maxPoint = entity.BoundingBox.MaxPoint;

                    // Log the coordinates
                    Console.WriteLine($"Entity Type: {entity.EntityType}");
                    Console.WriteLine($"MinPoint: ({minPoint.X}, {minPoint.Y}, {minPoint.Z})");
                    Console.WriteLine($"MaxPoint: ({maxPoint.X}, {maxPoint.Y}, {maxPoint.Z})");

                    // Get additional information about the entity
                    if (entity is BlockReferences blockRef)
                    {
                        Console.WriteLine($"Block Name: {blockRef.Name}");
                        Console.WriteLine($"Scale Factors: X={blockRef.XScaleFactor}, Y={blockRef.YScaleFactor}, Z={blockRef.ZScaleFactor}");
                        Console.WriteLine($"Rotation: {blockRef.Rotation}");
                    }

                    const double tolerance = 1e-6; // Adjust as needed

                    if (Math.Abs(minPoint.X - maxPoint.X) < tolerance || Math.Abs(minPoint.Y - maxPoint.Y) < tolerance)
                    {
                        Console.WriteLine("Rectangle has negligible width or height. Skipping.");
                    }
                    else
                    {
                        // Proceed to add the rectangle
                        AddRectangle(minPoint, maxPoint, newLayerName);
                        Console.WriteLine("Rectangle added successfully.");
                    }
                }
                else
                {
                    Console.WriteLine($"Entity '{entity.EntityType}' does not have a bounding box. Skipping.");
                }
            }

            int totalEntities = autoCADModelSpace.Count;
            Console.WriteLine($"Total entities in model space: {totalEntities}");
            Console.WriteLine($"Entities processed: {entities.Count}");

            Console.WriteLine("Finished adding bounding box rectangles.");
        }



        public void Dispose()
        {
            // Release entities
            if (entities != null)
            {
                foreach (var entity in entities)
                {
                    entity.Release();
                }
                entities.Clear();
                entities = null;
            }

            // Release layers
            if (layers != null)
            {

                foreach (var layer in layers)
                {
                    layer.Release();
                }

                layers.Clear();
                layers = null;
            }

            if (autoCADModelSpace != null)
            {
                _ = Marshal.ReleaseComObject(autoCADModelSpace);
                autoCADModelSpace = null;
            }
        }
    }
}
