using CADInteropServices.Objects.AutoCAD;
using CADInteropServices.Objects.MicroStation;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Orchestrators
{
	public class OrchestrateMicroStationExportDocumentData
	{
		public FileInfo microStationFile;

		public OrchestrateMicroStationExportDocumentData(
			string fileName)
		{
			microStationFile = new FileInfo(fileName);
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

			PerformMicroStationOperation(microStationDocument =>
			{
				FileInfo exportDataFile = new FileInfo(
					$"{microStationFile.DirectoryName}\\{microStationFile.Name}_exported_data.json");

				List<MicroStationElements> allElements = microStationDocument.PrepareElements();

				// Serialize entities to JSON
				string json = JsonConvert.SerializeObject(
					allElements,
					serializerSettings);

				// Write JSON to file
				File.WriteAllText(
					exportDataFile.FullName,
					json);

			});

		}


		private void PerformMicroStationOperation(
			Action<MicroStationDocuments> operation)
		{

			PerformMicroStationOperations microStationOperation = new PerformMicroStationOperations(
				microStationFile.FullName);

			microStationOperation.PerformMicroStationOperation(operation);


		}
	}
}
