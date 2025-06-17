using CADInteropServices.Orchestrators;

namespace TestCADInterOpServices
{
	[Apartment(ApartmentState.STA)]
	public class AutoCADInterOpTests
	{
		OrchestrateAutoCADExportDocumentData cadOrchestrator;

		[SetUp]
		public void Setup()
		{
			cadOrchestrator = new OrchestrateAutoCADExportDocumentData(
                "..\\..\\..\\samples\\sampleG.DWG");

		}

		[Test]
		public void TestExportAllData()
		{

			cadOrchestrator.ExportEntitiesToCsv();


			Assert.Pass();
		}

		[Test]
		public void TestReportAllLayers()
		{

			cadOrchestrator.ReportAllLayers();


			Assert.Pass();
		}

		[Test]
		public void TestReportEntityTypes()
		{

			cadOrchestrator.ReportAllEntityTypes();


			Assert.Pass();
		}

        [Test]
        public void TestReportBlocks()
        {

            cadOrchestrator.ReportAllReferenceBlockComponents();


            Assert.Pass();
        }

        [Test]
        public void PrintLayouts()
		{
			cadOrchestrator.ReportLayouts();
		}

        [Test]
        public void ExportToJson()
        {

			

            cadOrchestrator.ExportEntitiesToJson();


            Assert.Pass();
        }


    }
}