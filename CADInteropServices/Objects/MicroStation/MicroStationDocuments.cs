using System.Data;
using System.Runtime.InteropServices;
using Bentley.Interop.MicroStationDGN;

namespace CADInteropServices.Objects.MicroStation
{
    public class MicroStationDocuments : IDisposable
    {
        private DesignFile designFile;
        private ModelReferences models;
        private Application microStationApplication;

		public List<MicroStationModels> modelList;
        public FileInfo DesignFileInfo;

        public MicroStationDocuments()
        {
            // Default constructor if needed
        }

        public MicroStationDocuments(
            DesignFile designFile, 
            Application microStationApplication)
        {
            this.microStationApplication = microStationApplication;
            this.designFile = designFile;
            this.models = GetModels();
            this.modelList = GetModelList();
            this.DesignFileInfo = new FileInfo(designFile.FullName);

            
        }

        private ModelReferences GetModels()
        {
            return designFile.Models;
        }

        private List<MicroStationModels> GetModelList()
        {
            List<MicroStationModels> modelList = new List<MicroStationModels>();
            foreach (ModelReference modelRef in models)
            {
                MicroStationModels model = new MicroStationModels(modelRef, microStationApplication);
                modelList.Add(model);
            }
            return modelList;
        }

        // Data preparation methods
        public List<MicroStationElements> PrepareElements()
        {
            List<MicroStationElements> entities = new List<MicroStationElements>();

            foreach (MicroStationModels model in modelList)
            {
                List<MicroStationElements> modelElements = model.GetAllElements();
                entities.AddRange(modelElements);
            }

            return entities;
        }

        public DataTable PrepareEntitiesData()
        {
            DataTable entityTable = new DataTable("Entities");

            // Define columns
            entityTable.Columns.Add("ElementType", typeof(string));
            entityTable.Columns.Add("ID", typeof(long));
            entityTable.Columns.Add("ModelName", typeof(string));
            entityTable.Columns.Add("LevelName", typeof(string));
            entityTable.Columns.Add("Color", typeof(int));
            entityTable.Columns.Add("LineStyle", typeof(string));
            entityTable.Columns.Add("LineWeight", typeof(int));
            entityTable.Columns.Add("SpecificProperties", typeof(string));

            // Get all entities
            List<MicroStationElements> entities = PrepareElements();

            foreach (var entity in entities)
            {
                DataRow row = entityTable.NewRow();

                // Common properties
                row["ElementType"] = entity.ElementType;
                row["ID"] = entity.ElementID;
                row["ModelName"] = entity.ModelName;
                row["LevelName"] = entity.LevelName;
                row["Color"] = entity.Color;
                row["LineStyle"] = entity.LineStyle;
                row["LineWeight"] = entity.LineWeight;

                // Specific properties
                row["SpecificProperties"] = entity.GetSpecificPropertiesAsString();

                entityTable.Rows.Add(row);
            }

            return entityTable;
        }

        public DataTable PrepareLevelsData()
        {
            DataTable levelTable = new DataTable("Levels");

            // Define columns
            levelTable.Columns.Add("Name", typeof(string));
            levelTable.Columns.Add("IsOn", typeof(bool));
            levelTable.Columns.Add("Color", typeof(int));
            levelTable.Columns.Add("IsLocked", typeof(bool));

            // Keep track of processed levels to avoid duplicates
            HashSet<string> processedLevels = new HashSet<string>();

            foreach (MicroStationModels model in modelList)
            {
                List<MicroStationLevels> levels = model.GetAllLevels();

                foreach (var level in levels)
                {
                    if (!processedLevels.Contains(level.Name))
                    {
                        DataRow row = levelTable.NewRow();

                        row["Name"] = level.Name;
                        row["IsOn"] = level.IsOn;
                        row["Color"] = level.Color;
                        row["IsLocked"] = level.IsLocked;

                        levelTable.Rows.Add(row);
                        processedLevels.Add(level.Name);
                    }
                }
            }

            return levelTable;
        }

        // Data export methods
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
            string directoryName = null)
        {
            if (string.IsNullOrEmpty(directoryName))
            {
                directoryName = DesignFileInfo.DirectoryName;
            }

            DirectoryInfo directoryInfo = new DirectoryInfo(directoryName);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(DesignFileInfo.Name);
            string outputFilePathPrefix = $"{directoryInfo.FullName}\\{fileNameWithoutExtension}";

            // Prepare data
            DataTable entitiesTable = PrepareEntitiesData();
            FileInfo entitiesCsvFile = new FileInfo($"{outputFilePathPrefix}_entities_export.csv");
            ExportDataTableToCSV(entitiesTable, entitiesCsvFile.FullName);

            // Prepare levels data
            DataTable levelsTable = PrepareLevelsData();
            FileInfo levelsCsvFile = new FileInfo($"{outputFilePathPrefix}_levels_export.csv");
            ExportDataTableToCSV(levelsTable, levelsCsvFile.FullName);
        }

        // Save, SaveAs, Close, and Dispose methods
        public void Save()
        {
            designFile.Save();
        }

        public void SaveAs(
            string fileName)
        {
            try
            {
                designFile.SaveAs(fileName);
                Console.WriteLine("Design file saved as: " + fileName);
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
            if (modelList != null)
            {
                foreach (var model in modelList)
                {
                    model.Dispose();
                }
                modelList = null;
            }

            if (designFile != null)
            {
                designFile.Close();
                Marshal.ReleaseComObject(designFile);
                designFile = null;
            }
        }

        public void Dispose()
        {
            Close();
            GC.SuppressFinalize(this);
        }
    }


}
