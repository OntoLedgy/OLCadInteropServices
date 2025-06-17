using Newtonsoft.Json;
using System.Xml;
using System.Collections.Generic;
using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Objects.AutoCAD.Shapes;
using CADInteropServices.Objects.AutoCAD.Spaces;

namespace CADInteropServices.Orchestrators
{
    public class OrchestrateAutoCADExportDocumentData
	{
		public FileInfo autoCADFile;

		public OrchestrateAutoCADExportDocumentData(
			string fileName)
		{
			autoCADFile = new FileInfo(fileName);
		}

        public void ExportEntitiesToJson()
        {
            // Configure serialization settings
            var serializerSettings = new JsonSerializerSettings
            {
                Formatting = Newtonsoft.Json.Formatting.Indented, // For pretty-printing
                ReferenceLoopHandling = ReferenceLoopHandling.Ignore, // Ignore circular references
                NullValueHandling = NullValueHandling.Ignore, // Exclude null values
                TypeNameHandling = TypeNameHandling.Auto // Include type names
            };

            PerformAutoCADOperation(autoCADDocument =>
            {
                FileInfo exportDataFile = new FileInfo(
                    $"{autoCADFile.DirectoryName}\\{autoCADFile.Name}_exported_data.json");

                List<AutoCADEntities> allEntities = autoCADDocument.PrepareEntities();

                // Serialize entities to JSON
                string json = JsonConvert.SerializeObject(
                    allEntities, 
                    serializerSettings);

                // Write JSON to file
                File.WriteAllText(
                    exportDataFile.FullName, 
                    json);

            });


        }

        public void ExportEntitiesToCsv()
		{
			PerformAutoCADOperation(autoCADDocument =>
			{
				FileInfo exportDataFile = new FileInfo(
					$"{autoCADFile.DirectoryName}\\{autoCADFile.Name}_exported_data.csv");
				
				autoCADDocument.ExportAllData(
					exportDataFile.DirectoryName);
			});
		}


		public void ReportAllLayers()
		{
			PerformAutoCADOperation(autoCADDocument =>
			{
				autoCADDocument.ModelSpace.ReportAllLayers();
			});
		}

		public void ReportAllEntityTypes()
		{
			PerformAutoCADOperation(autoCADDocument =>
			{
				autoCADDocument.ModelSpace.ReportAllEntityTypes();
			});
		}

        public void ReportAllReferenceBlockComponents()
        {
            PerformAutoCADOperation(autoCADDocument =>
            {
                // Dictionary to store counts of entity types at each level
                Dictionary<int, Dictionary<string, int>> levelEntityTypeCounts = new Dictionary<int, Dictionary<string, int>>();

                // Initialize the visitedBlocks set
                HashSet<string> visitedBlocks = new HashSet<string>();

                // Get all entities from ModelSpace
                List<AutoCADEntities> allEntities = autoCADDocument.ModelSpace.GetAllEntities();

                foreach (var entity in allEntities)
                {
                    if (entity is BlockReferences blockRef)
                    {
                        // Start at level 1
                        CollectEntityTypes(
                            blockRef, 
                            levelEntityTypeCounts, 
                            1, 
                            visitedBlocks);
                    }
                }

                // Produce a summary report
                Console.WriteLine("Summary of Entity Types in Block References by Level:");

                foreach (var levelEntry in levelEntityTypeCounts.OrderBy(kvp => kvp.Key))
                {
                    int level = levelEntry.Key;
                    var entityTypeCounts = levelEntry.Value;

                    Console.WriteLine($"\nLevel {level}:");
                    foreach (var kvp in entityTypeCounts)
                    {
                        Console.WriteLine($"  {kvp.Key}: {kvp.Value}");
                    }
                }
            });
        }

        public void ReportLayouts()
        {
            PerformAutoCADOperation(autoCADDocument =>
            {
                List<Layouts> layouts = autoCADDocument.GetLayouts();

                foreach (Layouts layout in layouts)
                {
                    Console.WriteLine(layout.Name);
                }

            });

        }


        private void CollectEntityTypes(
            BlockReferences blockRef,
            Dictionary<int, Dictionary<string, int>> levelEntityTypeCounts,
            int level,
            HashSet<string> visitedBlocks)
        {
            Console.WriteLine($"Entering CollectEntityTypes for BlockReference '{blockRef.Name}' at level {level}");

            // Check if we've already visited this block reference
            if (visitedBlocks.Contains(blockRef.DefinitionHandle))
            {
                Console.WriteLine($"Already visited BlockReference '{blockRef.Name}' at level {level + 1}, skipping.");
                return;
            }

            // Add the current block definition's Handle to the visited set
            visitedBlocks.Add(blockRef.DefinitionHandle);

            // Initialize the dictionary for the current level if it doesn't exist
            if (!levelEntityTypeCounts.ContainsKey(level))
            {
                levelEntityTypeCounts[level] = new Dictionary<string, int>();
            }

            // Get the dictionary for the current level
            var entityTypeCounts = levelEntityTypeCounts[level];

            // Increment count for BlockReferences at current level
            if (entityTypeCounts.ContainsKey(blockRef.EntityType))
            {
                entityTypeCounts[blockRef.EntityType]++;
            }
            else
            {
                entityTypeCounts[blockRef.EntityType] = 1;
            }

            // Process entities within the block reference
            foreach (var transformedEntity in blockRef.TransformedEntities)
            {
                if (transformedEntity is BlockReferences nestedBlockRef)
                {
                    int nextLevel = level + 1; // Explicitly increment the level
                    Console.WriteLine($"Recursing into nested BlockReference '{nestedBlockRef.Name}' at level {nextLevel}");
                    // Recurse into nested block reference at nextLevel
                    CollectEntityTypes(nestedBlockRef, levelEntityTypeCounts, nextLevel, visitedBlocks);
                }
                else
                {
                    // Increment count for the entity type at the current level
                    if (entityTypeCounts.ContainsKey(transformedEntity.EntityType))
                    {
                        entityTypeCounts[transformedEntity.EntityType]++;
                    }
                    else
                    {
                        entityTypeCounts[transformedEntity.EntityType] = 1;
                    }

                    Console.WriteLine($"Processed entity '{transformedEntity.EntityType}' at level {level}");
                }
            }

            // Remove the current block definition's Handle from the visited set after processing
            //visitedBlocks.Remove(blockRef.DefinitionHandle);
        }




        private void PerformAutoCADOperation(
			Action<AutoCADDocuments> operation)
		{

            PerformAutoCADOperations autoCADOperations = new PerformAutoCADOperations(
                autoCADFile.FullName);

            autoCADOperations.PerformAutoCADOperation(operation);


		}
	}



	}
