using Autodesk.AutoCAD.Interop;
using System.Runtime.InteropServices;

namespace CADInteropServices.Objects.AutoCAD
{
    public class AutoCADApplications : IDisposable
    {

        public AcadApplication autoCADApplication = null;
        private static MessageFilter _messageFilter;

        public AutoCADApplications(
            bool visiblity)
        {
            try
            {
                // Ensure the thread is STA
                if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                {
                    throw new InvalidOperationException("The current thread must be set to STA (Single-Threaded Apartment) state.");
                }

                // Register the COM message filter
                _messageFilter = new MessageFilter();
                MessageFilter.Register();

                // Create a new instance of AutoCAD				
                autoCADApplication = new AcadApplication();
                autoCADApplication.Visible = visiblity;

            }
            catch (COMException comEx)
            {
                Console.WriteLine("COM Exception: " + comEx.Message);
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                throw;
            }
        }

        public AutoCADDocuments OpenDocument(
            string fileName)
        {

            // Wait for AutoCAD to initialize
            IAcadState acadState = autoCADApplication.GetAcadState();

            while (!acadState.IsQuiescent)
            {
                Thread.Sleep(500);
            }

            AcadDocument document = autoCADApplication.Documents.Open(
                fileName);

            // Wait for the document to be fully opened
            while (document == null)
            {
                Thread.Sleep(500);
            }

            AutoCADDocuments autoCADDocuments = new AutoCADDocuments(document);

            return autoCADDocuments;

        }


        public void Dispose()
        {
            // Clean up resources
            if (autoCADApplication != null)
            {
                autoCADApplication.Quit();
                Marshal.ReleaseComObject(autoCADApplication);
                autoCADApplication = null;
            }

            // Revoke the COM message filter
            MessageFilter.Revoke();

        }


    }
}
