using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Reflection.Metadata;
using CADInteropServices.Factories;
using CADInteropServices.Transformers;
using CADInteropServices.Objects.AutoCAD.Shapes;
using CADInteropServices.Objects.AutoCAD.Spaces;

namespace CADInteropServices.Objects.AutoCAD
{
    public class AutoCADDocuments : IDisposable
    {
        //TODO wrap this
        private AcadDocument? autoCADDocument;
        private AcadPaperSpace paperSpace;
        private AcadBlocks? blocks;

        private ModelSpaces? modelSpace;

        public List<Layouts> layouts;

        public FileInfo AutoCADFile;

        public AutoCADDocuments(
            AcadDocument document)
        {
            autoCADDocument = document;
            blocks = GetBlocks();
            modelSpace = GetModelSpace();
            paperSpace = GetPaperSpace();
            layouts = GetLayouts();
            AutoCADFile = new FileInfo(
                document.FullName);

        }

        private AcadBlocks GetBlocks()
        {
            return autoCADDocument.Blocks;
        }

        private ModelSpaces GetModelSpace()
        {
            AcadModelSpace modelSpace = null;

            const int maxRetries = 10;
            int retryCount = 0;
            bool success = false;

            while (!success && retryCount < maxRetries)
            {
                try
                {
                    modelSpace = autoCADDocument.ModelSpace;
                    success = true;
                }
                catch (COMException comEx) when ((uint)comEx.ErrorCode == 0x8001010A)
                {
                    retryCount++;
                    Console.WriteLine($"COM Exception: {comEx.Message}. Retrying {retryCount}/{maxRetries}...");
                    Thread.Sleep(500);
                }
            }

            if (!success || modelSpace == null)
            {
                Console.WriteLine("Failed to access ModelSpace after multiple retries.");
                return null;
            }

            return new ModelSpaces(modelSpace, blocks);
        }

        public ModelSpaces ModelSpace
        {
            get { return modelSpace; }
        }

        private AcadPaperSpace GetPaperSpace()
        {
            AcadPaperSpace paperSpace = null;

            const int maxRetries = 10;
            int retryCount = 0;
            bool success = false;

            while (!success && retryCount < maxRetries)
            {
                try
                {
                    paperSpace = autoCADDocument.PaperSpace;
                    success = true;
                }
                catch (COMException comEx) when ((uint)comEx.ErrorCode == 0x8001010A)
                {
                    retryCount++;
                    Console.WriteLine($"COM Exception: {comEx.Message}. Retrying {retryCount}/{maxRetries}...");
                    Thread.Sleep(500);
                }
            }

            if (!success || paperSpace == null)
            {
                Console.WriteLine("Failed to access PaperSpace after multiple retries.");
                return null;
            }

            return paperSpace;
        }

        public List<Layouts> GetLayouts()
        {

            try
            {
                List<Layouts> layoutlist = new List<Layouts>();

                AcadLayouts layouts = autoCADDocument.Layouts;

                foreach (AcadLayout autoCADLayout in layouts)
                {

                    Layouts layout = new Layouts(autoCADLayout, blocks);

                    layoutlist.Add(layout);


                }


                return layoutlist;

            }

            catch (COMException comEx)
            {
                Console.WriteLine($"COM Exception in PrintLayoutNames: {comEx.Message}");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception in PrintLayoutNames: {ex.Message}");
                return null;
            }


        }

        //data preparation

        public List<AutoCADEntities> PrepareEntities()
        {
            // Get all entities from ModelSpace
            List<AutoCADEntities> entities = modelSpace.GetAllEntities();

            // Process BlockReferences by passing the Blocks collection
            foreach (var entity in entities)
            {
                if (entity is BlockReferences blockRef)
                {
                    // The processing is done within the BlockReferences constructor
                    // Ensure that the Blocks collection is passed
                    blockRef.blocks = blocks;
                    // Alternatively, if blocks is readonly, consider modifying the constructor
                }
            }

            return entities;
        }

        public DataTable PrepareEntitiesData()
        {
            DataTable entityTable = new DataTable("Entities");

            // Define columns
            entityTable.Columns.Add("EntityType", typeof(string));
            entityTable.Columns.Add("Handle", typeof(string));
            entityTable.Columns.Add("Layer", typeof(string));
            entityTable.Columns.Add("Color", typeof(string));
            entityTable.Columns.Add("Linetype", typeof(string));
            entityTable.Columns.Add("Lineweight", typeof(double));

            // Add columns for specific properties as needed
            entityTable.Columns.Add("SpecificProperties", typeof(string));

            // Get all entities from ModelSpaces
            List<AutoCADEntities> entities = modelSpace.GetAllEntities();

            foreach (var entity in entities)
            {
                DataRow row = entityTable.NewRow();

                // Common properties
                row["EntityType"] = entity.EntityType;
                row["Handle"] = entity.Handle;
                row["Layer"] = entity.Layer;
                row["Color"] = entity.Color;
                row["Linetype"] = entity.Linetype;
                row["Lineweight"] = entity.Lineweight;

                // Specific properties
                row["SpecificProperties"] = entity.GetSpecificPropertiesAsString();

                entityTable.Rows.Add(row);
            }

            return entityTable;
        }

        public DataTable PrepareLayersData()
        {
            DataTable layerTable = new DataTable("Layers");

            // Define columns
            layerTable.Columns.Add("Name", typeof(string));
            layerTable.Columns.Add("IsOn", typeof(bool));
            layerTable.Columns.Add("IsFrozen", typeof(bool));
            layerTable.Columns.Add("IsLocked", typeof(bool));
            layerTable.Columns.Add("ColourName", typeof(string));
            layerTable.Columns.Add("Red", typeof(int));
            layerTable.Columns.Add("Green", typeof(int));
            layerTable.Columns.Add("Blue", typeof(int));

            // Get all layers from ModelSpaces
            List<AutoCADDocumentLayers> layers = modelSpace.GetAllLayers();

            foreach (var layer in layers)
            {
                DataRow row = layerTable.NewRow();

                row["Name"] = layer.Name;
                row["IsOn"] = layer.IsOn;
                row["IsFrozen"] = layer.IsFrozen;
                row["IsLocked"] = layer.IsLocked;
                row["ColourName"] = layer.ColourName;
                row["Red"] = layer.TrueColour.Red;
                row["Green"] = layer.TrueColour.Green;
                row["Blue"] = layer.TrueColour.Blue;

                layerTable.Rows.Add(row);
            }

            return layerTable;
        }

        //data export

        public void ExportDataTableToCSV(
            DataTable table,
            string filePath)
        {

            using (StreamWriter writer = new StreamWriter(filePath))
            {
                // Write column headers
                IEnumerable<string> columnNames = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
                writer.WriteLine(string.Join(",", columnNames));

                // Write rows
                foreach (DataRow row in table.Rows)
                {
                    IEnumerable<string> fields = row.ItemArray.Select(field =>
                    {
                        if (field == null)
                            return "";

                        string fieldString = field.ToString();

                        // Escape fields containing commas or quotes
                        if (fieldString.Contains(",") || fieldString.Contains("\""))
                        {
                            fieldString = $"\"{fieldString.Replace("\"", "\"\"")}\"";
                        }

                        return fieldString;
                    });

                    writer.WriteLine(string.Join(",", fields));
                }
            }
        }

        public void ExportAllData(
            string? directoryName = null)
        {

            if (string.IsNullOrEmpty(directoryName))
            {
                directoryName = AutoCADFile.DirectoryName;
            }

            DirectoryInfo directoryInfo = new DirectoryInfo(directoryName);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(AutoCADFile.Name);
            string outputFilePathPrefix = $"{directoryInfo.FullName}\\{fileNameWithoutExtension}";
            // Prepare data
            DataTable entitiesTable = PrepareEntitiesData();

            FileInfo entitiesCsvFile = new($"{outputFilePathPrefix}_entities_export.csv");

            ExportDataTableToCSV(entitiesTable, entitiesCsvFile.FullName);

            // Similarly for layers
            DataTable layersTable = PrepareLayersData();

            FileInfo layersCsvFile = new($"{outputFilePathPrefix}_layers_export.csv");

            ExportDataTableToCSV(layersTable, layersCsvFile.FullName);

        }


        public void Save()
        {
            autoCADDocument.Save();
        }

        public void SaveAs(string fileName)
        {
            try
            {
                autoCADDocument.SaveAs(fileName);
                Console.WriteLine("Document saved as: " + fileName);
            }
            catch (COMException comEx)
            {
                Console.WriteLine("COM Exception in SaveAs: " + comEx.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception in SaveAs: " + ex.Message);
            }
        }

        public void Close()
        {
            if (modelSpace != null)
            {
                modelSpace.Dispose();
                modelSpace = null;
            }

            if (blocks != null)
            {
                Marshal.ReleaseComObject(blocks);
                blocks = null;
            }

            if (autoCADDocument != null)
            {
                autoCADDocument.Close();
                Marshal.ReleaseComObject(autoCADDocument);
                autoCADDocument = null;
            }
        }

        public void Dispose()
        {
            Close();
            GC.SuppressFinalize(this);
        }
    }
}
