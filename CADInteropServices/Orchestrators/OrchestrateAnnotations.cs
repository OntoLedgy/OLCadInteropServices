using CADInteropServices.Objects.AutoCAD;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Orchestrators
{

    public class OrchestrateAnnotations
	{
		FileInfo autoCADFile { get; set; }

		public OrchestrateAnnotations(
						string fileName)
		{
			autoCADFile = new FileInfo(fileName);
		}

		public void AddLayers()
		{
            PerformAutoCADOperation(autoCADDocument =>
            {
                autoCADDocument.ModelSpace.AddLayer("Annotations");
            });

        }


		public void AnnotateAllEntities()
		{		

			PerformAutoCADOperation(autoCADDocument =>
			{
				autoCADDocument.ModelSpace.AddBoundingBoxRectangles("Annotations");
			});
		}


		private void PerformAutoCADOperation(
			Action<AutoCADDocuments> operation)
		{
            string directory = autoCADFile.DirectoryName;
            string filenameWithoutExtension = Path.GetFileNameWithoutExtension(autoCADFile.Name);
            string extension = autoCADFile.Extension;
            string targetFileName = Path.Combine(directory, filenameWithoutExtension + "_annotated" + extension);


            PerformAutoCADOperations autoCADOperations = new PerformAutoCADOperations(
				autoCADFile.FullName);

			autoCADOperations.PerformAutoCADOperation(operation, targetFileName);


		}

	}
}
