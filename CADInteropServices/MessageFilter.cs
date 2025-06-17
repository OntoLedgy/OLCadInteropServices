using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CADInteropServices
{
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

	public class MessageFilter : IOleMessageFilter
	{
		// Register the message filter
		public static void Register()
		{
			IOleMessageFilter newFilter = new MessageFilter();

			IOleMessageFilter oldFilter = null;

			CoRegisterMessageFilter(
				newFilter, 
				out oldFilter);

		}

		// Revoke the message filter
		public static void Revoke()
		{
			IOleMessageFilter oldFilter = null;
			CoRegisterMessageFilter(null, out oldFilter);
		}

		// Handle incoming thread requests
		int IOleMessageFilter.HandleInComingCall(
			int dwCallType, 
			IntPtr hTaskCaller, 
			int dwTickCount, 
			IntPtr lpInterfaceInfo)
		{
			return 0; // SERVERCALL_ISHANDLED
		}

		// Retry rejected calls
		int IOleMessageFilter.RetryRejectedCall(
			IntPtr hTaskCallee, 
			int dwTickCount, 
			int dwRejectType)
		{
			if (dwRejectType == 2) // SERVERCALL_RETRYLATER
			{
				// Retry the call immediately
				return 500;
			}
			return -1; // Cancel the call
		}

		// Message pending
		int IOleMessageFilter.MessagePending(
			IntPtr hTaskCallee, 
			int dwTickCount, 
			int dwPendingType)
		{
			return 2; // PENDINGMSG_WAITDEFPROCESS
		}

		// P/Invoke declarations
		[DllImport("Ole32.dll")]
		private static extern int CoRegisterMessageFilter(
			IOleMessageFilter newFilter, 
			out IOleMessageFilter oldFilter);
	}


}
