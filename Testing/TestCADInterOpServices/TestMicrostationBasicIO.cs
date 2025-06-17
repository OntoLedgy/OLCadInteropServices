using CADInteropServices.Orchestrators;

namespace TestCADInterOpServices
{
	[Apartment(ApartmentState.STA)]
	public class MicrostationInterOpTests
	{
		OrchestrateMicroStationExportDocumentData cadOrchestrator;

		[SetUp]
		public void Setup()
		{
			cadOrchestrator = new OrchestrateMicroStationExportDocumentData(
                "..\\..\\..\\samples\\sampleK.DGN");

		}

		[Test]
		public void TestExportToJson()
		{

			cadOrchestrator.ExportEntitiesToJson();


			Assert.Pass();
		}




    }
}