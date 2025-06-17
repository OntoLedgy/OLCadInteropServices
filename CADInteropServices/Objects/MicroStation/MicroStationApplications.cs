using System.Runtime.InteropServices;
using Bentley.Interop.MicroStationDGN;


namespace CADInteropServices.Objects.MicroStation
{
    public class MicroStationApplications : IDisposable
    {
        public Bentley.Interop.MicroStationDGN.Application microStationApplication = null;
        private static MessageFilter _messageFilter;

        public MicroStationApplications(bool visibility)
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

                // Create a new instance of MicroStation
                Type microStationType = Type.GetTypeFromProgID("MicroStationDGN.Application");
                microStationApplication = (Bentley.Interop.MicroStationDGN.Application)Activator.CreateInstance(microStationType);
                microStationApplication.Visible = visibility;
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

        public MicroStationDocuments OpenDocument(string fileName)
        {
            // Wait for MicroStation to initialize (if necessary)
            while (microStationApplication == null)
            {
                Thread.Sleep(500);
            }

            // Open the design file
            DesignFile designFile = microStationApplication.OpenDesignFile(fileName, ReadOnly: false);

            // Wait for the design file to be fully opened
            while (designFile == null)
            {
                Thread.Sleep(500);
            }

			MicroStationDocuments microStationDocuments = new MicroStationDocuments(designFile, microStationApplication);

            return microStationDocuments;
        }

        public void Dispose()
        {
            // Clean up resources
            if (microStationApplication != null)
            {
                microStationApplication.Quit();
                Marshal.ReleaseComObject(microStationApplication);
                microStationApplication = null;
            }

            // Revoke the COM message filter
            MessageFilter.Revoke();
        }
    }


    // MessageFilter class to handle COM threading issues
    public class MessageFilter : IOleMessageFilter
    {
        // Start the filter.
        public static void Register()
        {
            IOleMessageFilter newFilter = new MessageFilter();
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(newFilter, out oldFilter);
        }

        // Done with the filter, close it.
        public static void Revoke()
        {
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(null, out oldFilter);
        }

        // IOleMessageFilter functions.
        int IOleMessageFilter.HandleInComingCall(int dwCallType,
          System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr lpInterfaceInfo)
        {
            // Return the flag SERVERCALL_ISHANDLED.
            return 0;
        }

        int IOleMessageFilter.RetryRejectedCall(System.IntPtr hTaskCallee,
          int dwTickCount, int dwRejectType)
        {
            if (dwRejectType == 2)
            {
                // Retry the thread call immediately if return >=0 & <100.
                return 99;
            }
            // Too busy; cancel call.
            return -1;
        }

        int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee,
          int dwTickCount, int dwPendingType)
        {
            // Return the flag PENDINGMSG_WAITDEFPROCESS.
            return 2;
        }

        // Implement the IOleMessageFilter interface.
        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter,
            out IOleMessageFilter oldFilter);
    }

    [ComImport(), Guid("00000016-0000-0000-C000-000000000046"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(
            int dwCallType,
            IntPtr hTaskCaller,
            int dwTickCount,
            IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwRejectType);

        [PreserveSig]
        int MessagePending(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwPendingType);
    }
}
