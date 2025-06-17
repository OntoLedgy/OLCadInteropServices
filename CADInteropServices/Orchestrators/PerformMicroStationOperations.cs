using CADInteropServices.Objects.MicroStation;
using System.Runtime.InteropServices;


namespace CADInteropServices.Orchestrators
{
    public class PerformMicroStationOperations
	{
		public FileInfo microStationFile;

		public PerformMicroStationOperations(
						string fileName)
		{
			microStationFile = new FileInfo(fileName);
		}

		public void PerformMicroStationOperation(
				Action<MicroStationDocuments> operation,
				string saveFileName = null)
		{
			try
			{
				// Ensure the thread is STA
				if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
				{
					throw new Exception("The current thread must be set to STA (Single-Threaded Apartment) state.");
				}

				MicroStationApplications microStationApplication = new MicroStationApplications(
					false);

				MicroStationDocuments microStationDocument = microStationApplication.OpenDocument(
					microStationFile.FullName);

				Console.WriteLine("Reading DGN file: " + microStationDocument.DesignFileInfo.FullName);


				operation(
					microStationDocument);

				if (saveFileName != null)
				{
					microStationDocument.SaveAs(saveFileName);
					Console.WriteLine("Document saved as: " + saveFileName);
				}

				microStationDocument.Close();
				microStationApplication.Dispose();
			}
			catch (COMException comEx)
			{
				Console.WriteLine("COM Exception: " + comEx.Message);
				Console.WriteLine("Error Code: " + comEx.ErrorCode);
				Console.WriteLine("Exception Details: " + comEx.ToString());
			}
			catch (Exception ex)
			{
				Console.WriteLine("Exception: " + ex.Message);
				Console.WriteLine("Exception Details: " + ex.ToString());
			}
		}
	}
}
