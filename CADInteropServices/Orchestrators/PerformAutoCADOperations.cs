using CADInteropServices.Objects.AutoCAD;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices.Orchestrators
{
    public class PerformAutoCADOperations
	{
		public FileInfo autoCADFile;

		public PerformAutoCADOperations(
						string fileName)
		{
			autoCADFile = new FileInfo(fileName);
		}

		public void PerformAutoCADOperation(
				Action<AutoCADDocuments> operation,
				string saveFileName = null)
		{
			try
			{
				// Ensure the thread is STA
				if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
				{
					throw new Exception("The current thread must be set to STA (Single-Threaded Apartment) state.");
				}

				AutoCADApplications autoCADApplication = new AutoCADApplications(
					false);

				AutoCADDocuments autoCADDocument = autoCADApplication.OpenDocument(
					autoCADFile.FullName);

				Console.WriteLine("Reading DWG file: " + autoCADDocument.AutoCADFile.FullName);


				operation(
					autoCADDocument);

				if (saveFileName != null)
				{
                    autoCADDocument.SaveAs(saveFileName);
                    Console.WriteLine("Document saved as: " + saveFileName);
                }

				autoCADDocument.Close();
				autoCADApplication.Dispose();
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
