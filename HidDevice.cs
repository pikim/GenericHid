using Microsoft.Win32.SafeHandles;
using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Timers;
using System.Windows.Forms;

namespace GenericHid
{
   ///<summary>
   ///
   /// Project: GenericHid
   ///
   /// ***********************************************************************
   /// Software License Agreement
   ///
   /// Licensor grants any person obtaining a copy of this software ("You")
   /// a worldwide, royalty-free, non-exclusive license, for the duration of
   /// the copyright, free of charge, to store and execute the Software in a
   /// computer system and to incorporate the Software or any portion of it
   /// in computer programs You write.
   ///
   /// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
   /// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
   /// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
   /// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
   /// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
   /// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
   /// THE SOFTWARE.
   /// ***********************************************************************
   ///
   /// Author
   /// Jan Axelson
   ///
   /// This software was written using Visual Studio Express 2012 for Windows
   /// Desktop building for the .NET Framework v4.5.
   ///
   /// Purpose:
   /// Demonstrates USB communications with a generic HID-class device
   ///
   /// Requirements:
   /// Windows Vista or later and an attached USB generic Human Interface Device (HID).
   /// (Does not run on Windows XP or earlier because .NET Framework 4.5 will not install on these OSes.)
   ///
   /// Description:
   /// Finds an attached device that matches the vendor and product IDs in the form's
   /// text boxes.
   ///
   /// Retrieves the device's capabilities.
   /// Sends and requests HID reports.
   ///
   /// Uses the System.Management class and Windows Management Instrumentation (WMI) to detect
   /// when a device is attached or removed.
   ///
   /// A list box displays the data sent and received along with error and status messages.
   /// You can select data to send and 1-time or periodic transfers.
   ///
   /// You can change the size of the host's Input report buffer and request to use control
   /// transfers only to exchange Input and Output reports.
   ///
   /// To view additional debugging messages, in the Visual Studio development environment,
   /// from the main menu, select Build > Configuration Manager > Active Solution Configuration
   /// and select Configuration > Debug and from the main menu, select View > Output.
   ///
   /// The application uses asynchronous FileStreams to read Input reports and write Output
   /// reports so the application's main thread doesn't have to wait for the device to retrieve a
   /// report when the HID driver's buffer is empty or send a report when the device's endpoint is busy.
   ///
   /// For code that finds a device and opens handles to it, see the FindTheHid routine in frmMain.cs.
   /// For code that reads from the device, see GetInputReportViaInterruptTransfer,
   /// GetInputReportViaControlTransfer, and GetFeatureReport in Hid.cs.
   /// For code that writes to the device, see SendInputReportViaInterruptTransfer,
   /// SendInputReportViaControlTransfer, and SendFeatureReport in Hid.cs.
   ///
   /// This project includes the following modules:
   ///
   /// GenericHid.cs - runs the application.
   /// FrmMain.cs - routines specific to the form.
   /// Hid.cs - routines specific to HID communications.
   /// DeviceManagement.cs - routine for obtaining a handle to a device from its GUID.
   /// Debugging.cs - contains a routine for displaying API error messages.
   /// HidDeclarations.cs - Declarations for API functions used by Hid.cs.
   /// FileIODeclarations.cs - Declarations for file-related API functions.
   /// DeviceManagementDeclarations.cs - Declarations for API functions used by DeviceManagement.cs.
   /// DebuggingDeclarations.cs - Declarations for API functions used by Debugging.cs.
   ///
   /// Companion device firmware for several device CPUs is available from www.Lvr.com/hidpage.htm
   /// You can use any generic HID (not a system mouse or keyboard) that sends and receives reports.
   /// This application will not detect or communicate with non-HID-class devices.
   ///
   /// For more information about HIDs and USB, and additional example device firmware to use
   /// with this application, visit Lakeview Research at http://Lvr.com
   /// Send comments, bug reports, etc. to jan@Lvr.com or post on my PORTS forum: http://www.lvr.com/forum
   ///
   /// V6.2
   /// 11/12/13
   /// Disabled form buttons when a transfer is in progress.
   /// Other minor edits for clarity and readability.
   /// Will NOT run on Windows XP or earlier, see below.
   ///
   /// V6.1
   /// 10/28/13
   /// Uses the .NET System.Management class to detect device arrival and removal with WMI instead of Win32 RegisterDeviceNotification.
   /// Other minor edits.
   /// Will NOT run on Windows XP or earlier, see below.
   ///
   /// V6.0
   /// 2/8/13
   /// This version will NOT run on Windows XP or earlier because the code uses .NET Framework 4.5 to support asynchronous FileStreams.
   /// The .NET Framework 4.5 redistributable is compatible with Windows 8, Windows 7 SP1, Windows Server 2008 R2 SP1,
   /// Windows Server 2008 SP2, Windows Vista SP2, and Windows Vista SP3.
   /// For compatibility, replaced ToInt32 with ToInt64 here:
   /// IntPtr pDevicePathName = new IntPtr(detailDataBuffer.ToInt64() + 4);
   /// and here:
   /// if ((deviceNotificationHandle.ToInt64() == IntPtr.Zero.ToInt64()))
   /// For compatibility if the charset isn't English, added System.Globalization.CultureInfo.InvariantCulture here:
   /// if ((String.Compare(DeviceNameString, mydevicePathName, true, System.Globalization.CultureInfo.InvariantCulture) == 0))
   /// Replaced all Microsoft.VisualBasic namespace code with other .NET equivalents.
   /// Revised user interface for more flexibility.
   /// Moved interrupt-transfer and other HID-specific code to Hid.cs.
   /// Used JetBrains ReSharper to clean up the code: http://www.jetbrains.com/resharper/
   ///
   /// V5.0
   /// 3/30/11
   /// Replaced ReadFile and WriteFile with FileStreams. Thanks to Joe Dunne and John on my Ports forum for tips on this.
   /// Simplified Hid.cs.
   /// Replaced the form timer with a system timer.
   ///
   /// V4.6
   /// 1/12/10
   /// Supports Vendor IDs and Product IDs up to FFFFh.
   ///
   /// V4.52
   /// 11/10/09
   /// Changed HIDD_ATTRIBUTES to use UInt16
   ///
   /// V4.51
   /// 2/11/09
   /// Moved Free_ and similar to Finally blocks to ensure they execute.
   ///
   /// V4.5
   /// 2/9/09
   /// Changes to support 64-bit systems, memory management, and other corrections.
   /// Big thanks to Peter Nielsen.
   ///
   /// </summary>

   public sealed partial class HidDevice
   {
      //  Used in error messages.

      private Hid myHid;
      private Int16 myVendorID;
      private Int16 myProductID;
      private String myManufacturerString;
      private String myProductString;
      private String mySerialNumber;
      private String myDevicePathName;
      private Boolean myDeviceDetected;
      private Boolean myDeviceHandlesObtained;
      private SafeFileHandle readHandle;
      private SafeFileHandle writeHandle;
      private FileStream fsDeviceDataRead;
      private FileStream fsDeviceDataWrite;
      private Boolean controlTransferInProgress;
      private Boolean interruptTransferInProgress;
      private Boolean receivePermanently;
      private TransferTypes transferType;
      private Double readTimeout;
      private static System.Timers.Timer tmrReadTimeout;
      private DeviceManagement myDeviceManagement;

      //  For viewing results of API calls via Debug.Write.
//    private Debugging myDebugging = new Debugging();
      private static readonly TraceSource traceSource = new TraceSource("HidDevice");
      private Boolean showMsgBoxOnException;

      private ManagementEventWatcher deviceArrivedWatcher;
      private ManagementEventWatcher deviceRemovedWatcher;
      public event DeviceNotifyDelegate Inserted;
      public event DeviceNotifyDelegate Removed;
      public event DeviceNotifyDelegate ReadTimedOut;
      public event DataReceiveDelegate AsynchDataReceived;


      ///  <summary>
      ///  Creates a HidDevice with the given Vendor ID, Product ID, manufacturer string,
      ///  product string and serial number string.
      ///  </summary>
      ///
      ///  <param name="manufacturer">The manufacturer string.</param>
      ///  <param name="product">The product string.</param>
      ///  <param name="serialNumber">The serial number string.</param>
      ///
      ///  <remarks>
      ///  vID and pID are obligatory, while the other parameters also accept null or an empty string if not all
      ///  strings are needed. If only a part of a string is interesting, only that part needs to be passed.
      ///  </remarks>

      public HidDevice(Int16 vID, Int16 pID, String manufacturer, String product, String serialNumber) : this(vID, pID)
      {
         // Strings must not be null because of program logic
         if (manufacturer != null) myManufacturerString = manufacturer;
         if (product != null) myProductString = product;
         if (serialNumber != null) mySerialNumber = serialNumber;
      }

      ///  <summary>
      ///  Creates a HidDevice with the given Vendor ID and Product ID.
      ///  </summary>

      public HidDevice(Int16 vID, Int16 pID)
      {
         //  Set USB Vendor ID and Product ID:

         myVendorID = vID;
         myProductID = pID;
         myManufacturerString = String.Empty;
         myProductString = String.Empty;
         mySerialNumber = String.Empty;
         myDeviceManagement = new DeviceManagement();

         controlTransferInProgress = false;
         interruptTransferInProgress = false;
         receivePermanently = false;
         transferType = TransferTypes.Control;
         showMsgBoxOnException = false;

         readTimeout = 30000;
         tmrReadTimeout = new System.Timers.Timer(readTimeout);
         tmrReadTimeout.Elapsed += new ElapsedEventHandler(OnReadTimeout);
         tmrReadTimeout.Stop();
      }

      ///  <summary>
      ///  Provides a central mechanism for exception handling.
      ///  Displays a message box that describes the exception.
      ///  </summary>
      ///
      ///  <param name="moduleName">The module where the exception occurred.</param>
      ///  <param name="e">The exception.</param>
      ///  <param name="showMsgBox">Determine if the exception shall also be displayed in a message box.</param>

      internal static void DisplayException(String moduleName, Exception e, Boolean showMsgBox)
      {
         //  Create error message.

         if (showMsgBox)
         {
            const String caption = "Unexpected Exception";

            String message = "Exception: " + e.Message.TrimEnd('\r', '\n') + Environment.NewLine +
                             "Module: " + moduleName + Environment.NewLine +
                             "Method: " + e.TargetSite.Name;

            MessageBox.Show(message, caption, MessageBoxButtons.OK);
         }

         traceSource.TraceEvent(TraceEventType.Error, 1, " Exception in {0}: {1}{2}{3}{2}   {4}",
                               moduleName, e.Message.TrimEnd('\r', '\n'), Environment.NewLine,
                               e.StackTrace, e.TargetSite);
//         traceSource.TraceData(TraceEventType.Error, 1, e);

         // Get the last error and display it.

         Int32 error = Marshal.GetLastWin32Error();

         traceSource.TraceEvent(TraceEventType.Error, 1, " The last Win32 Error was: {0}", error);
      }

      ///  <summary>
      ///  Perform actions that must execute when the program starts.
      ///  </summary>

      public void Open()
      {
         try
         {
            myHid = new Hid();

            tmrReadTimeout.Stop();

            traceSource.TraceEvent(TraceEventType.Information, 0, "------------------------------------------------------------");
            traceSource.TraceEvent(TraceEventType.Information, 1, "GenericHID loaded for device VID=0x{0:X04} & PID=0x{1:X04}.",
                                   myVendorID, myProductID);

            DeviceNotificationsStart();
            myDeviceDetected = FindDeviceUsingWmi();
            if (myDeviceDetected && !myDeviceHandlesObtained)
            {
               myDeviceHandlesObtained = FindTheHid();
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }
      }

      ///  <summary>
      ///  Add handlers to detect device arrival and removal.
      ///  </summary>

      private void DeviceNotificationsStart()
      {
         AddDeviceArrivedHandler();
         AddDeviceRemovedHandler();
      }

      ///  <summary>
      ///  Stop receiving notifications about device arrival and removal.
      ///  </summary>

      private void DeviceNotificationsStop()
      {
         try
         {
            if (deviceArrivedWatcher != null)
               deviceArrivedWatcher.Stop();
            if (deviceRemovedWatcher != null)
               deviceRemovedWatcher.Stop();
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }
      }

      ///  <summary>
      ///  Add a handler to detect arrival of devices using WMI.
      ///  </summary>

      private void AddDeviceArrivedHandler()
      {
         const Int32 pollingIntervalSeconds = 3;
         ManagementScope scope = new ManagementScope("root\\CIMV2");
         scope.Options.EnablePrivileges = true;

         try
         {
            WqlEventQuery q = new WqlEventQuery();
            q.EventClassName = "__InstanceCreationEvent";
            q.WithinInterval = new TimeSpan(0, 0, pollingIntervalSeconds);
            q.Condition = @"TargetInstance ISA 'Win32_USBControllerDevice'";

            deviceArrivedWatcher = new ManagementEventWatcher(scope, q);
            deviceArrivedWatcher.EventArrived += DeviceAdded;
            deviceArrivedWatcher.Start();
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
            if (deviceArrivedWatcher != null)
               deviceArrivedWatcher.Stop();
         }
      }

      ///  <summary>
      ///  Add a handler to detect removal of devices using WMI.
      ///  </summary>

      private void AddDeviceRemovedHandler()
      {
         const Int32 pollingIntervalSeconds = 3;
         ManagementScope scope = new ManagementScope("root\\CIMV2");
         scope.Options.EnablePrivileges = true;

         try
         {
            WqlEventQuery q = new WqlEventQuery();
            q.EventClassName = "__InstanceDeletionEvent";
            q.WithinInterval = new TimeSpan(0, 0, pollingIntervalSeconds);
            q.Condition = @"TargetInstance ISA 'Win32_USBControllerdevice'";

            deviceRemovedWatcher = new ManagementEventWatcher(scope, q);
            deviceRemovedWatcher.EventArrived += DeviceRemoved;
            deviceRemovedWatcher.Start();
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
            if (deviceRemovedWatcher != null)
               deviceRemovedWatcher.Stop();
         }
      }

      ///  <summary>
      ///  Called on arrival of any device.
      ///  Calls a routine that searches to see if the desired device is present.
      ///  </summary>

      private void DeviceAdded(object sender, EventArrivedEventArgs e)
      {
         traceSource.TraceEvent(TraceEventType.Verbose, 1, "A USB device has been inserted");

         try
         {
            if (!myDeviceDetected || !myDeviceHandlesObtained)
            {
               myDeviceDetected = FindDeviceUsingWmi();

               if (myDeviceDetected)
               {
                  myDeviceHandlesObtained = FindTheHid();

                  if (myDeviceHandlesObtained)
                  {
                     if (Inserted != null) Inserted.Invoke();

                     if (receivePermanently)
                     {
                        ReadInputReport();
                     }
                  }
               }
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }
      }

      ///  <summary>
      ///  Called on removal of any device.
      ///  Calls a routine that searches to see if the desired device is still present.
      ///  </summary>

      private void DeviceRemoved(object sender, EventArrivedEventArgs e)
      {
         traceSource.TraceEvent(TraceEventType.Verbose, 1, "A USB device has been removed");

         try
         {
            if (myDeviceDetected || myDeviceHandlesObtained)
            {
               myDeviceDetected = FindDeviceUsingWmi();

               if (!myDeviceDetected)
               {
                  CloseCommunications();

                  if (Removed != null) Removed.Invoke();
               }
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }
      }

      ///  <summary>
      ///  Use the System.Management class to find a device by Vendor ID and Product ID using WMI. If found, display device properties.
      ///  </summary>
      ///
      ///  <remarks>
      ///  During debugging, if you stop the firmware but leave the device attached, the device may still be detected as present
      ///  but will be unable to communicate. The device will show up in Windows Device Manager as well.
      ///  This situation is unlikely to occur with a final product.
      ///  </remarks>

      private Boolean FindDeviceUsingWmi()
      {
         Boolean deviceDetected = false;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "FindDeviceUsingWmi");

         try
         {
            // Prepend "@" to string below to treat backslash as a normal character (not escape character):

            String deviceIdString = @"USB\VID_" + myVendorID.ToString("X4") + "&PID_" + myProductID.ToString("X4");

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PnPEntity");

            foreach (ManagementObject queryObj in searcher.Get())
            {
               if (queryObj["PNPDeviceID"].ToString().Contains(deviceIdString)
                && queryObj["PNPDeviceID"].ToString().Contains(mySerialNumber))
               {
                  deviceDetected = true;
                  traceSource.TraceEvent(TraceEventType.Information, 1, " My device found (WMI)");

                  // Display device properties.

                  foreach (WmiDeviceProperties wmiDeviceProperty in Enum.GetValues(typeof(WmiDeviceProperties)))
                  {
                     traceSource.TraceEvent(TraceEventType.Verbose, 1, "  {0}: {1}", wmiDeviceProperty, queryObj[wmiDeviceProperty.ToString()]);
                  }
               }
            }

            if (!deviceDetected)
            {
               traceSource.TraceEvent(TraceEventType.Information, 1, " My device not found (WMI)");
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }

         return deviceDetected;
     }

      ///  <summary>
      ///  Uses a series of API calls to locate a HID - class device
      ///  by its Vendor ID and Product ID.
      ///  </summary>
      ///
      ///  <returns>True if the device is detected, False if not detected.</returns>

      private Boolean FindTheHid()
      {
         Boolean deviceHandleObtained = false;
         Boolean devicesFound = false;
         Boolean hidFound = false;
         Boolean success = false;
         Guid hidGuid = Guid.Empty;
         Int32 memberIndex = 0;
         SafeFileHandle hidHandle = null;
         String[] devicePathName = new String[128];
         String hidUsage;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "FindTheHid");

         try
         {
            CloseCommunications();

            // Get the HID-class GUID.

            hidGuid = myHid.GetHidGuid();

            traceSource.TraceEvent(TraceEventType.Verbose, 1, " GUID for system HIDs: {0}", hidGuid);

            //  Fill an array with the device path names of all attached HIDs.

            devicesFound = myDeviceManagement.FindDeviceFromGuid(hidGuid, ref devicePathName);

            //  If there is at least one HID, attempt to read the Vendor ID and Product ID
            //  of each device until there is a match or all devices have been examined.

            if (devicesFound)
            {
               memberIndex = 0;

               do
               {
                  //  ***
                  //  API function:
                  //  CreateFile

                  //  Purpose:
                  //  Retrieves a handle to a device.

                  //  Accepts:
                  //  A device path name returned by SetupDiGetDeviceInterfaceDetail
                  //  The type of access requested (read/write).
                  //  FILE_SHARE attributes to allow other processes to access the device while this handle is open.
                  //  A Security structure or IntPtr.Zero.
                  //  A creation disposition value. Use OPEN_EXISTING for devices.
                  //  Flags and attributes for files. Not used for devices.
                  //  Handle to a template file. Not used.

                  //  Returns: a handle without read or write access.
                  //  This enables obtaining information about all HIDs, even system
                  //  keyboards and mice.
                  //  Separate handles are used for reading and writing.
                  //  ***

                  //  Open the handle without read/write access to enable getting information about any HID, even system keyboards and mice.

                  hidHandle = FileIo.CreateFile(devicePathName[memberIndex], FileIo.AccessNone, FileIo.FileShareRead | FileIo.FileShareWrite, IntPtr.Zero, FileIo.OpenExisting, 0, IntPtr.Zero);

                  if (!hidHandle.IsInvalid)
                  {
                     //  The returned handle is valid,
                     //  so find out if this is the device we're looking for.

                     //  Set the Size property of DeviceAttributes to the number of bytes in the structure.

                     myHid.DeviceAttributes.Size = Marshal.SizeOf(myHid.DeviceAttributes);

                     success = myHid.GetAttributes(hidHandle, ref myHid.DeviceAttributes);

                     if (success)
                     {
                        traceSource.TraceEvent(TraceEventType.Verbose, 1, " HIDD_ATTRIBUTES structure filled without error.");
                        traceSource.TraceEvent(TraceEventType.Verbose, 1, "  Structure size: {0}", myHid.DeviceAttributes.Size);
                        traceSource.TraceEvent(TraceEventType.Verbose, 1, "  Vendor ID: 0x{0:X04}", myHid.DeviceAttributes.VendorID);
                        traceSource.TraceEvent(TraceEventType.Verbose, 1, "  Product ID: 0x{0:X04}", myHid.DeviceAttributes.ProductID);
                        traceSource.TraceEvent(TraceEventType.Verbose, 1, "  Version Number: 0x{0:X}", myHid.DeviceAttributes.VersionNumber);

                        //  Find out if the device matches the one we're looking for.

                        if ((myHid.DeviceAttributes.VendorID == myVendorID) && (myHid.DeviceAttributes.ProductID == myProductID))
                        {
                           //  Read the serial number

                           String manufacturer = myHid.GetManufacturerString(hidHandle);
                           String product = myHid.GetProductString(hidHandle);
                           String serialNumber = myHid.GetSerialNumberString(hidHandle);

                           //  Display the information in form's list box.

                           traceSource.TraceEvent(TraceEventType.Information, 1, " Device detected");
                           traceSource.TraceEvent(TraceEventType.Verbose, 1, "  Vendor ID: 0x{0:X04}", myHid.DeviceAttributes.VendorID);
                           traceSource.TraceEvent(TraceEventType.Verbose, 1, "  Product ID: 0x{0:X04}", myHid.DeviceAttributes.ProductID);
                           traceSource.TraceEvent(TraceEventType.Verbose, 1, "  Manufacturer: {0}", manufacturer);
                           traceSource.TraceEvent(TraceEventType.Verbose, 1, "  Product: {0}", product);
                           traceSource.TraceEvent(TraceEventType.Verbose, 1, "  Serial number: {0}", serialNumber);

                           if (manufacturer.Contains(myManufacturerString) &&
                               product.Contains(myProductString) &&
                               serialNumber.Contains(mySerialNumber))
                           {
                              //  Either the descriptor string(s) match or don't matter

                              hidFound = true;
  
                              //  Save the DevicePathName for OnDeviceChange().
  
                              myDevicePathName = devicePathName[memberIndex];
                           }
                        }
                        else
                        {
                           //  It's not a match, so close the handle.

                           hidHandle.Close();
                        }
                     }
                     else
                     {
                        //  There was a problem in retrieving the information.

                        traceSource.TraceEvent(TraceEventType.Verbose, 1, " Error in filling HIDD_ATTRIBUTES structure.");
                        hidHandle.Close();
                     }
                  }

                  //  Keep looking until we find the device or there are no devices left to examine.

                  memberIndex = memberIndex + 1;
               }
               while ( ! (hidFound || (memberIndex == devicePathName.Length)));
            }

            if (hidFound)
            {
               //  The device was detected.
               //  Register to receive notifications if the device is removed or attached.

               //  Learn the capabilities of the device.

               myHid.Capabilities = myHid.GetDeviceCapabilities(hidHandle);

               //  Find out if the device is a system mouse or keyboard.

               hidUsage = myHid.GetHidUsage(myHid.Capabilities);

               //  Get the Input report buffer size.

               GetInputReportBufferSize(hidHandle);

               //  Close the handle and reopen it with read/write access.

               hidHandle.Close();
               readHandle = FileIo.CreateFile(myDevicePathName, FileIo.GenericRead, FileIo.FileShareRead | FileIo.FileShareWrite, IntPtr.Zero, FileIo.OpenExisting, FileIo.FileFlagOverlapped, IntPtr.Zero);
               writeHandle = FileIo.CreateFile(myDevicePathName, FileIo.GenericWrite, FileIo.FileShareRead | FileIo.FileShareWrite, IntPtr.Zero, FileIo.OpenExisting, FileIo.FileFlagOverlapped, IntPtr.Zero);

               if (readHandle.IsInvalid || writeHandle.IsInvalid)
               {
                  traceSource.TraceEvent(TraceEventType.Verbose, 1, " The device is a system {0}.", hidUsage);
                  traceSource.TraceEvent(TraceEventType.Verbose, 1, " Windows 2000 and Windows XP obtain exclusive access to Input and Output reports for this devices.");
                  traceSource.TraceEvent(TraceEventType.Verbose, 1, " Applications can access Feature reports only.");
               }
               else
               {
                  if (myHid.Capabilities.InputReportByteLength > 0)
                  {
                     //  Set the size of the Input report buffer.

                     fsDeviceDataRead = new FileStream(readHandle, FileAccess.Read, myHid.Capabilities.InputReportByteLength, true);
                  }

                  if (myHid.Capabilities.OutputReportByteLength > 0)
                  {
                     //  Set the size of the Input report buffer.

                     fsDeviceDataWrite = new FileStream(writeHandle, FileAccess.Write, myHid.Capabilities.OutputReportByteLength, true);
                  }

                  //  Flush any waiting reports in the buffers. (optional)

                  myHid.FlushQueue(readHandle);
                  myHid.FlushQueue(writeHandle);

                  deviceHandleObtained = true;
               }
            }
            else
            {
               //  The device wasn't detected.

               traceSource.TraceEvent(TraceEventType.Information, 1, " My device not detected.");
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }

         return deviceHandleObtained;
     }

      ///  <summary>
      ///  Displays received or written report data.
      ///  </summary>
      ///
      ///  <param name="buffer">Contains the report data.</param>
      ///  <param name="currentReportType" >"Input", "Output", or "Feature".</param>
      ///  <param name="currentReadOrWritten" >"read" for Input and IN Feature reports, "written" for Output and OUT Feature reports.</param>

      private void DisplayReportData(HidReport report, ReportTypes currentReportType, ReportAction currentReadOrWritten)
      {
         String byteValue = String.Empty;
         Int32 count;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, " DisplayReportData");

         try
         {
            traceSource.TraceEvent(TraceEventType.Verbose, 1, " {0} report has been {1}.", currentReportType, currentReadOrWritten);

            //  Display the report data received in the form's list box.

            traceSource.TraceEvent(TraceEventType.Verbose, 1, "  Report ID: 0x{0:X2}", report.ReportId);

            for (count = 0; count < report.Data.Length; count++)
            {
               //  Display bytes as 2-character Hex strings.

               byteValue += String.Format(" 0x{0:X2}", report.Data[count]);
            }
            traceSource.TraceEvent(TraceEventType.Verbose, 1, "  Report Data:{0}", byteValue);
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }
      }

      ///  <summary>
      ///  Retrieves a Feature report.
      ///  </summary>
      ///
      ///  <param name="report">A report containing the ID and the datalength of the report to be read.</param>
      ///
      ///  <returns>HidReport if read was successful, null if not.</returns>

      public HidReport ReadFeatureReport(HidReport report)
      {
         return ReadFeatureReport(report.ReportId, report.Length);
      }

      ///  <summary>
      ///  Retrieves a Feature report.
      ///  </summary>
      ///
      ///  <param name="reportID">The ID of the report to be read.</param>
      ///
      ///  <returns>HidReport if read was successful, null if not.</returns>

      public HidReport ReadFeatureReport(Byte reportID)
      {
         return ReadFeatureReport(reportID, myHid.Capabilities.FeatureReportByteLength);
      }

      ///  <summary>
      ///  Retrieves a Feature report.
      ///  </summary>
      ///
      ///  <param name="reportID">The ID of the report to be read.</param>
      ///  <param name="dataLength">The number of bytes to be read (including the report ID).</param>
      ///
      ///  <returns>HidReport if read was successful, null if not.</returns>

      public HidReport ReadFeatureReport(Byte reportID, Int16 dataLength)
      {
         HidReport featureReport = null;
         SafeFileHandle hidHandle;
         Byte[] inFeatureReportBuffer = null;
         Boolean success = false;

         //  Report header for the debug display:
         traceSource.TraceEvent(TraceEventType.Verbose, 1, "***** HID Read Feature Report (ID=0x{0}) *****",
                                String.Format("{0:X02}", reportID));

         try
         {
            //  If the device has been detected before but timed out or failed on a
            //  previous attempt to access it, look for the device.

            if (myDeviceDetected && !myDeviceHandlesObtained)
            {
               myDeviceHandlesObtained = FindTheHid();
            }

            if (myDeviceHandlesObtained)
            {
               if (myHid.Capabilities.FeatureReportByteLength > 0)
               {
                  //  The HID has a Feature report.

                  if (controlTransferInProgress)
                  {
                     return null;
                  }
                  controlTransferInProgress = true;

                  //  Set the size of the Feature report buffer.
                  //  Subtract 1 from the value in the Capabilities structure because
                  //  the array begins at index 0.

                  inFeatureReportBuffer = new Byte[dataLength];

                  //  Store the report ID in the buffer:

                  inFeatureReportBuffer[0] = reportID;

                  //  Read a report from the device.

                  hidHandle = FileIo.CreateFile(myDevicePathName, FileIo.AccessNone, FileIo.FileShareRead | FileIo.FileShareWrite, IntPtr.Zero, FileIo.OpenExisting, 0, IntPtr.Zero);
                  success = myHid.GetFeatureReport(hidHandle, ref inFeatureReportBuffer);
                  hidHandle.Close();
                  controlTransferInProgress = false;

                  if (success)
                  {
                     featureReport = new HidReport(inFeatureReportBuffer);

                     DisplayReportData(featureReport, ReportTypes.Feature, ReportAction.Read);
                  }
                  else
                  {
                     CloseCommunications();
                     traceSource.TraceEvent(TraceEventType.Verbose, 1, " The attempt to read a Feature report failed.");
                  }
               }
               else
               {
                  traceSource.TraceEvent(TraceEventType.Verbose, 1, " The HID doesn't have a Feature report.");
               }
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }

         return featureReport;
      }

      ///  <summary>
      ///  Sends a Feature report.
      ///  </summary>
      ///
      ///  <param name="featureReport">The HidReport to be written.</param>

      public Boolean WriteFeatureReport(HidReport featureReport)
      {
         SafeFileHandle hidHandle;
         Byte[] outFeatureReportBuffer = null;
         Boolean success = false;

         //  Report header for the debug display:
         traceSource.TraceEvent(TraceEventType.Verbose, 1, "***** HID Write Feature Report (ID=0x{0}) *****",
                                String.Format("{0:X02}", featureReport.ReportId));

         try
         {
            //  If the device has been detected before but timed out or failed on a
            //  previous attempt to access it, look for the device.

            if (myDeviceDetected && !myDeviceHandlesObtained)
            {
               myDeviceHandlesObtained = FindTheHid();
            }

            if (myDeviceHandlesObtained)
            {
               if (myHid.Capabilities.FeatureReportByteLength > 0)
               {
                  //  The HID has a Feature report.

                  if (controlTransferInProgress)
                  {
                     return false;
                  }
                  controlTransferInProgress = true;

                  //  Set the size of the Feature report buffer.
                  //  Subtract 1 from the value in the Capabilities structure because
                  //  the array begins at index 0.

                  outFeatureReportBuffer = new Byte[myHid.Capabilities.FeatureReportByteLength];

                  //  Store the report ID in the buffer and the report data following the report ID.

                  outFeatureReportBuffer = featureReport.GetBytes(outFeatureReportBuffer.Length);

                  //  Write a report to the device

                  hidHandle = FileIo.CreateFile(myDevicePathName, FileIo.AccessNone, FileIo.FileShareRead | FileIo.FileShareWrite, IntPtr.Zero, FileIo.OpenExisting, 0, IntPtr.Zero);
                  success = myHid.SendFeatureReport(hidHandle, outFeatureReportBuffer);
                  hidHandle.Close();
                  controlTransferInProgress = false;

                  if (success)
                  {
                     DisplayReportData(featureReport, ReportTypes.Feature, ReportAction.Written);
                  }
                  else
                  {
                     CloseCommunications();
                     traceSource.TraceEvent(TraceEventType.Verbose, 1, "The attempt to send a Feature report failed.");
                  }
               }
               else
               {
                  traceSource.TraceEvent(TraceEventType.Verbose, 1, "The HID doesn't have a Feature report.");
               }
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }

         return success;
      }

      ///  <summary>
      ///  Retrieves an Input report using the default TransferType.
      ///  </summary>
      ///
      ///  <returns>HidReport if read was successful, null if not.</returns>

      public HidReport ReadInputReport()
      {
        return ReadInputReport(transferType);
      }

      ///  <summary>
      ///  Retrieves an Input report.
      ///  </summary>
      ///
      ///  <param name="controlTransfer">Select whether a control transfer shall be used or a FileStream.</param>
      ///
      ///  <returns>HidReport if read was successful, null if not.</returns>

      public HidReport ReadInputReport(TransferTypes transferType)
      {
         Byte[] inputReportBuffer = null;
         HidReport inputReport = null;
         Boolean success = false;

         //  Report header for the debug display:
         traceSource.TraceEvent(TraceEventType.Verbose, 1, "***** HID Read Report *****");

         try
         {
            //  If the device has been detected before but timed out or failed on a
            //  previous attempt to access it, look for the device.

            if (myDeviceDetected && !myDeviceHandlesObtained)
            {
               myDeviceHandlesObtained = FindTheHid();
            }

            if (myDeviceHandlesObtained)
            {
               //  Don't attempt to exchange reports if valid handles aren't available
               //  (as for a mouse or keyboard under Windows 2000/XP.)

               if ((readHandle != null) && (!(readHandle.IsInvalid)))
               {
                  //  Read an Input report.

                  //  Don't attempt to send an Input report if the HID has no Input report.
                  //  (The HID spec requires all HIDs to have an interrupt IN endpoint,
                  //  which suggests that all HIDs must support Input reports.)

                  if (myHid.Capabilities.InputReportByteLength > 0)
                  {
                     //  Set the size of the Input report buffer.

                     inputReportBuffer = new Byte[myHid.Capabilities.InputReportByteLength];

                     if (transferType.Equals(TransferTypes.Control))
                     {
                        if (controlTransferInProgress)
                        {
                           return null;
                        }
                        controlTransferInProgress = true;

                        //  Read a report using a control transfer.

                        success = myHid.GetInputReportViaControlTransfer(readHandle, ref inputReportBuffer);
                        controlTransferInProgress = false;

                        if (success)
                        {
                           inputReport = new HidReport(inputReportBuffer);

                           DisplayReportData(inputReport, ReportTypes.Input, ReportAction.Read);
                        }
                        else
                        {
                           CloseCommunications();
                           traceSource.TraceEvent(TraceEventType.Verbose, 1, " The attempt to read an Input report has failed.");
                        }
                     }
                     else
                     {
                        //  Read a report using interrupt transfers.
                        //  To enable reading a report without blocking the main thread, this
                        //  application uses an asynchronous delegate.

                        if (fsDeviceDataRead.CanRead)
                        {
                           if (interruptTransferInProgress)
                           {
                              return null;
                           }
                           interruptTransferInProgress = true;

                           if (readTimeout > 0)
                           {
                              //  Timeout if no report is available.

                              tmrReadTimeout.Start();
                           }

                           fsDeviceDataRead.BeginRead(inputReportBuffer, 0, inputReportBuffer.Length, new AsyncCallback(OnInputReportReadComplete), inputReportBuffer);
                        }
                        else
                        {
                           CloseCommunications();
                           traceSource.TraceEvent(TraceEventType.Verbose, 1, " The attempt to read an async Input report has failed.");
                        }
                     }
                  }
                  else
                  {
                     traceSource.TraceEvent(TraceEventType.Verbose, 1, " The HID doesn't have an Input report.");
                  }
               }
               else
               {
                  traceSource.TraceEvent(TraceEventType.Verbose, 1, " Invalid handle.");
               }
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }

         return inputReport;
      }

      ///  <summary>
      ///  Sends an Output report using the default TransferType.
      ///  </summary>
      ///
      ///  <param name="outputReport">The HidReport to be written.</param>

      public Boolean WriteOutputReport(HidReport outputReport)
      {
        return WriteOutputReport(outputReport, transferType);
      }

      ///  <summary>
      ///  Sends an Output report.
      ///  </summary>
      ///
      ///  <param name="outputReport">The HidReport to be written.</param>
      ///  <param name="controlTransfer">Select whether a control transfer shall be used or a FileStream.</param>

      public Boolean WriteOutputReport(HidReport outputReport, TransferTypes transferType)
      {
         Byte[] outputReportBuffer = null;
         Boolean success = false;

         //  Report header for the debug display:
         traceSource.TraceEvent(TraceEventType.Verbose, 1, "***** HID Write Report *****");

         try
         {
            //  If the device has been detected before but timed out or failed on a
            //  previous attempt to access it, look for the device.

            if (myDeviceDetected && !myDeviceHandlesObtained)
            {
               myDeviceHandlesObtained = FindTheHid();
            }

            if (myDeviceHandlesObtained)
            {
               //  Don't attempt to exchange reports if valid handles aren't available
               //  (as for a mouse or keyboard under Windows 2000/XP.)

               if ((writeHandle != null) && (!(writeHandle.IsInvalid)))
               {
                  //  Write an Output report.

                  //  Don't attempt to send an Output report if the HID has no Output report.

                  if (myHid.Capabilities.OutputReportByteLength > 0)
                  {
                     //  Set the size of the Output report buffer.

                     outputReportBuffer = new Byte[myHid.Capabilities.OutputReportByteLength];

                     //  Store the report ID in the buffer and the report data following the report ID.

                     outputReportBuffer = outputReport.GetBytes(outputReportBuffer.Length);

                     //  Write a report.

                     if (transferType.Equals(TransferTypes.Control))
                     {
                        if (controlTransferInProgress)
                        {
                           return false;
                        }
                        controlTransferInProgress = true;

                        //  Use a control transfer to send the report,
                        //  even if the HID has an interrupt OUT endpoint.

                        success = myHid.SendOutputReportViaControlTransfer(writeHandle, outputReportBuffer);
                        controlTransferInProgress = false;
                     }
                     else
                     {
                        //  If the HID has an interrupt OUT endpoint, the host uses an
                        //  interrupt transfer to send the report.
                        //  If not, the host uses a control transfer.

                        if (fsDeviceDataWrite.CanWrite)
                        {
                           fsDeviceDataWrite.Write(outputReportBuffer, 0, outputReportBuffer.Length);
                           success = true;
                        }
                     }

                     if (success)
                     {
                        DisplayReportData(outputReport, ReportTypes.Output, ReportAction.Written);
                     }
                     else
                     {
                        CloseCommunications();
                        traceSource.TraceEvent(TraceEventType.Verbose, 1, " The attempt to write an Output report failed.");
                     }
                  }
                  else
                  {
                     traceSource.TraceEvent(TraceEventType.Verbose, 1, " The HID doesn't have an Output report.");
                  }
               }
               else
               {
                  traceSource.TraceEvent(TraceEventType.Verbose, 1, " Invalid handle.");
               }
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }

         return success;
      }

      ///  <summary>
      ///  Finds and returns the manufacturer string of the device.
      ///  </summary>
      ///
      ///  <returns>The manufacturer string or an empty string if an error occurred.</returns>

      public String GetManufacturerString()
      {
         SafeFileHandle hidHandle;
         Boolean success = false;
         String manufacturerString = "";

         try
         {
            hidHandle = FileIo.CreateFile(myDevicePathName, FileIo.AccessNone, FileIo.FileShareRead | FileIo.FileShareWrite, IntPtr.Zero, FileIo.OpenExisting, 0, IntPtr.Zero);
            manufacturerString = myHid.GetManufacturerString(hidHandle);
            hidHandle.Close();
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }

         return manufacturerString;
      }

      ///  <summary>
      ///  Finds and returns the product string of the device.
      ///  </summary>
      ///
      ///  <returns>The product string or an empty string if an error occurred.</returns>

      public String GetProductString()
      {
         SafeFileHandle hidHandle;
         Boolean success = false;
         String productString = "";

         try
         {
            hidHandle = FileIo.CreateFile(myDevicePathName, FileIo.AccessNone, FileIo.FileShareRead | FileIo.FileShareWrite, IntPtr.Zero, FileIo.OpenExisting, 0, IntPtr.Zero);
            productString = myHid.GetProductString(hidHandle);
            hidHandle.Close();
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }

         return productString;
      }

      ///  <summary>
      ///  Finds and returns the serial number string of the device.
      ///  </summary>
      ///
      ///  <returns>The serial number string or an empty string if an error occurred.</returns>

      public String GetSerialNumberString()
      {
         SafeFileHandle hidHandle;
         Boolean success = false;
         String serialNumberString = "";

         try
         {
            hidHandle = FileIo.CreateFile(myDevicePathName, FileIo.AccessNone, FileIo.FileShareRead | FileIo.FileShareWrite, IntPtr.Zero, FileIo.OpenExisting, 0, IntPtr.Zero);
            serialNumberString = myHid.GetSerialNumberString(hidHandle);
            hidHandle.Close();
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }

         return serialNumberString;
      }

      ///  <summary>
      ///  Finds and returns the number of Input buffers
      ///  (the number of Input reports the host will store).
      ///  </summary>
      ///
      ///  <returns>The size of the Input buffer.</returns>

      public Int32 GetInputReportBufferSize()
      {
         Int32 bufferSize = 0;

         try
         {
            bufferSize = GetInputReportBufferSize(readHandle);
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }

         return bufferSize;
      }

      private Int32 GetInputReportBufferSize(SafeFileHandle handle)
      {
         Int32 numberOfInputBuffers = 0;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "GetInputReportBufferSize");

         try
         {
            //  Get the number of input buffers.

            myHid.GetNumberOfInputBuffers(handle, ref numberOfInputBuffers);

            traceSource.TraceEvent(TraceEventType.Verbose, 1, " Number of Input buffers: {0}", numberOfInputBuffers);
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }

         return numberOfInputBuffers;
      }

      ///  <summary>
      ///  Set the number of Input buffers (the number of Input reports
      ///  the host will store) from the value in the text box.
      ///  </summary>
      ///
      ///  <param name="numberOfInputBuffers">The desired Input buffer size.</param>
      ///
      ///  <returns>True if the size was changed successfully, False if not.</returns>

      public Boolean SetInputReportBufferSize(Int32 numberOfInputBuffers)
      {
         Int32 resultingNumberOfBuffers = 0;
         Boolean success = false;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "SetInputReportBufferSize");

         try
         {
            if (controlTransferInProgress || interruptTransferInProgress)
            {
               return false;
            }

            //  Set the number of buffers.

            myHid.SetNumberOfInputBuffers(readHandle, numberOfInputBuffers);

            traceSource.TraceEvent(TraceEventType.Verbose, 1, " Number of Input buffers: {0}", numberOfInputBuffers);

            //  Verify and display the result.

            resultingNumberOfBuffers = GetInputReportBufferSize();

            if (resultingNumberOfBuffers == numberOfInputBuffers)
            {
               success = true;
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }

         return success;
      }

      ///  <summary>
      ///  Retrieves Input report data and status information.
      ///  This routine is called automatically when myInputReport.Read
      ///  returns. Calls several marshaling routines to access the main form.
      ///  </summary>
      ///
      ///  <param name="ar">An object containing status information about
      ///  the asynchronous operation.</param>

      private void OnInputReportReadComplete(IAsyncResult ar)
      {
         Int32 bytesRead;
         HidReport inputReport = null;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "OnInputReportReadComplete");

         try
         {
            if (ar.IsCompleted)
            {
               bytesRead = fsDeviceDataRead.EndRead(ar);

               tmrReadTimeout.Stop();

               inputReport = new HidReport((Byte[])ar.AsyncState);

               //  Generate event and pass data.
               if (AsynchDataReceived != null) AsynchDataReceived.Invoke(inputReport);

               DisplayReportData(inputReport, ReportTypes.Input, ReportAction.Read);
            }
            else
            {
               CloseCommunications();
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " The attempt to read an async Input report has failed");
            }
         }
         catch (Exception ex)
         {
            CloseCommunications();
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }
         finally
         {
            //  Enable requesting another transfer.

            interruptTransferInProgress = false;

            if (myDeviceHandlesObtained && receivePermanently)
            {
               ReadInputReport();
            }
         }
      }

      ///  <summary>
      ///  System timer timeout if read via interrupt transfer doesn't return.
      ///  </summary>
      ///
      ///  <param name="source"></param>
      ///  <param name="e"></param>

      private void OnReadTimeout(object source, ElapsedEventArgs e)
      {
         traceSource.TraceEvent(TraceEventType.Verbose, 1, "OnReadTimeout");

         try
         {
//            CloseCommunications();

            tmrReadTimeout.Stop();

            //  Generate event and pass data.
            if (ReadTimedOut != null) ReadTimedOut.Invoke();
         }
         finally
         {
            //  Enable requesting another transfer.

            interruptTransferInProgress = false;

            if (receivePermanently)
            {
               ReadInputReport();
            }
         }
      }

      ///  <summary>
      ///  Close the handle and FileStreams for a device.
      ///  </summary>

      private void CloseCommunications()
      {
         traceSource.TraceEvent(TraceEventType.Verbose, 1, " CloseCommunications");

         if (fsDeviceDataRead != null)
         {
            fsDeviceDataRead.Close();
         }

         if (fsDeviceDataWrite != null)
         {
            fsDeviceDataWrite.Close();
         }

         if ((readHandle != null) && (!(readHandle.IsInvalid)))
         {
            readHandle.Close();
         }

         if ((writeHandle != null) && (!(writeHandle.IsInvalid)))
         {
            writeHandle.Close();
         }

         //  The next attempt to communicate will get new handles and FileStreams.

         myDeviceHandlesObtained = false;
         controlTransferInProgress = false;
         interruptTransferInProgress = false;
      }

      ///  <summary>
      ///  Perform actions that must execute when the program ends.
      ///  </summary>

      public void Close()
      {
         traceSource.TraceEvent(TraceEventType.Verbose, 1, "Closing Device");

         try
         {
            CloseCommunications();

            DeviceNotificationsStop();

            traceSource.Close();
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }
      }

      ///  <summary>
      ///  Finalize method.
      ///  </summary>

      ~HidDevice()
      {
         try
         {
            Close();
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, showMsgBoxOnException);
         }
      }

      ///  <summary>
      ///  Returns true if the device was found and handles were retrieved successfully.
      ///  </summary>

      public Boolean IsDetected
      {
         get {return myDeviceHandlesObtained;}
      }

      ///  <summary>
      ///  Enables or disables a popup dialog when an exception occurs.
      ///  Returns the state of this setting.
      ///  </summary>

      public Boolean ShowMsgBoxOnException
      {
         get {return showMsgBoxOnException;}
         set
         {
            showMsgBoxOnException = value;
            myHid.ShowMsgBoxOnException = value;
            myDeviceManagement.ShowMsgBoxOnException = value;
         }
      }

      ///  <summary>
      ///  Enables or disables permanent reading from the HID device.
      ///  Returns the state of this setting.
      ///  </summary>

      public Boolean ReceivePermanently
      {
         get {return receivePermanently;}
         set {receivePermanently = value;}
      }

      ///  <summary>
      ///  Returns true if there is currently a transfer in progress.
      ///  </summary>

      public Boolean TransferInProgress
      {
         get {return (controlTransferInProgress || interruptTransferInProgress);}
      }

      ///  <summary>
      ///  Sets the read timeout duration in milliseconds.
      ///  Returns the state of this setting.
      ///  </summary>

      public Double ReadTimeout
      {
         get {return tmrReadTimeout.Interval;}
         set
         {
            Boolean tmrState = tmrReadTimeout.Enabled;
            tmrReadTimeout.Enabled = false;
            readTimeout = value;
            if (readTimeout > 0)
            {
               tmrReadTimeout.Interval = readTimeout;
               tmrReadTimeout.Enabled = tmrState;
            }
         }
      }

      ///  <summary>
      ///  Sets the default transfer type in case no type is specified (Control or Interrupt).
      ///  Returns the state of this setting.
      ///  </summary>

      public TransferTypes TransferType
      {
         get {return transferType;}
         set {transferType = value;}
      }
   }
}
