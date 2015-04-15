using Microsoft.Win32.SafeHandles;
using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace GenericHid
{
   ///  <summary>
   ///  Supports Windows API functions for accessing HID-class USB devices.
   ///  Includes routines for retrieving information about the configuring a HID and
   ///  sending and receiving reports via control and interrupt transfers.
   ///  </summary>

   internal sealed partial class Hid
   {
      //  Used in error messages.

      internal NativeMethods.HIDP_CAPS Capabilities;
      internal NativeMethods.HIDD_ATTRIBUTES DeviceAttributes;

      //  For viewing results of API calls via Debug.Write.
//      internal static Debugging myDebugging = new Debugging();
      private static readonly TraceSource traceSource = new TraceSource("Hid");
      public Boolean ShowMsgBoxOnException
      {
         get; set;
      }

      ///  <summary>
      ///  Provides a central mechanism for exception handling.
      ///  Displays a message box that describes the exception.
      ///  </summary>
      ///
      ///  <param name="moduleName">The module where the exception occurred.</param>
      ///  <param name="e">The exception.</param>

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
      ///  Remove any input reports waiting in the buffer.
      ///  </summary>
      ///
      ///  <param name="hidHandle">A handle to a device.</param>
      ///
      ///  <returns>True on success. False on failure.</returns>

      internal Boolean FlushQueue(SafeFileHandle hidHandle)
      {
         Boolean success = false;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "FlushQueue");

         try
         {
            //  ***
            //  API function: HidD_FlushQueue

            //  Purpose: Removes any Input reports waiting in the buffer.

            //  Accepts: a handle to the device.

            //  Returns: True on success, False on failure.
            //  ***

            success = NativeMethods.HidD_FlushQueue(hidHandle);
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return success;
      }

      ///  <summary>
      ///  Get the HID-class GUID.
      ///  </summary>
      ///
      ///  <returns>The GUID.</returns>

      internal Guid GetHidGuid()
      {
         Guid hidGuid = Guid.Empty;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "GetHidGuid");

         try
         {
            //  ***
            //  API function: 'HidD_GetHidGuid

            //  Purpose: Retrieves the interface class GUID for the HID class.

            //  Accepts: A System.Guid object for storing the GUID.
            //  ***

            NativeMethods.HidD_GetHidGuid(ref hidGuid);
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return hidGuid;
      }

      ///  <summary>
      ///  Get HID attributes.
      ///  </summary>
      ///
      ///  <param name="hidHandle">HID handle retrieved with CreateFile.</param>
      ///  <param name="deviceAttributes">HID attributes structure.</param>
      ///
      ///  <returns>True on success. False on failure.</returns>

      internal Boolean GetAttributes(SafeFileHandle hidHandle, ref NativeMethods.HIDD_ATTRIBUTES deviceAttributes)
      {
         Boolean success = false;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "GetAttributes");

         try
         {
            //  ***
            //  API function:
            //  HidD_GetAttributes

            //  Purpose:
            //  Retrieves a HIDD_ATTRIBUTES structure containing the Vendor ID,
            //  Product ID, and Product Version Number for a device.

            //  Accepts:
            //  A handle returned by CreateFile.
            //  A pointer to receive a HIDD_ATTRIBUTES structure.

            //  Returns:
            //  True on success, False on failure.
            //  ***

            success = NativeMethods.HidD_GetAttributes(hidHandle, ref deviceAttributes);
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return success;
      }

      ///  <summary>
      ///  Retrieves a structure with information about a device's capabilities.
      ///  </summary>
      ///
      ///  <param name="hidHandle">A handle to a device.</param>
      ///
      ///  <returns>A HIDP_CAPS structure.</returns>

      internal NativeMethods.HIDP_CAPS GetDeviceCapabilities(SafeFileHandle hidHandle)
      {
         var preparsedData = new IntPtr();

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "GetDeviceCapabilities");

         try
         {
            //  ***
            //  API function: HidD_GetPreparsedData

            //  Purpose: retrieves a pointer to a buffer containing information about the device's capabilities.
            //  HidP_GetCaps and other API functions require a pointer to the buffer.

            //  Requires:
            //  A handle returned by CreateFile.
            //  A pointer to a buffer.

            //  Returns:
            //  True on success, False on failure.
            //  ***

            NativeMethods.HidD_GetPreparsedData(hidHandle, ref preparsedData);

            //  ***
            //  API function: HidP_GetCaps

            //  Purpose: find out a device's capabilities.
            //  For standard devices such as joysticks, you can find out the specific
            //  capabilities of the device.
            //  For a custom device where the software knows what the device is capable of,
            //  this call may be unneeded.

            //  Accepts:
            //  A pointer returned by HidD_GetPreparsedData
            //  A pointer to a HIDP_CAPS structure.

            //  Returns: True on success, False on failure.
            //  ***

            Int32 result = NativeMethods.HidP_GetCaps(preparsedData, ref Capabilities);
            if ((result != 0))
            {
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Usage: {0:X}", Capabilities.Usage);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Usage Page: {0:X}", Capabilities.UsagePage);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Input Report Byte Length: {0}", Capabilities.InputReportByteLength);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Output Report Byte Length: {0}", Capabilities.OutputReportByteLength);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Feature Report Byte Length: {0}", Capabilities.FeatureReportByteLength);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Number of Link Collection Nodes: {0}", Capabilities.NumberLinkCollectionNodes);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Number of Input Button Caps: {0}", Capabilities.NumberInputButtonCaps);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Number of Input Value Caps: {0}", Capabilities.NumberInputValueCaps);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Number of Input Data Indices: {0}", Capabilities.NumberInputDataIndices);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Number of Output Button Caps: {0}", Capabilities.NumberOutputButtonCaps);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Number of Output Value Caps: {0}", Capabilities.NumberOutputValueCaps);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Number of Output Data Indices: {0}", Capabilities.NumberOutputDataIndices);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Number of Feature Button Caps: {0}", Capabilities.NumberFeatureButtonCaps);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Number of Feature Value Caps: {0}", Capabilities.NumberFeatureValueCaps);
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " Number of Feature Data Indices: {0}", Capabilities.NumberFeatureDataIndices);

               //  ***
               //  API function: HidP_GetValueCaps

               //  Purpose: retrieves a buffer containing an array of HidP_ValueCaps structures.
               //  Each structure defines the capabilities of one value.
               //  This application doesn't use this data.

               //  Accepts:
               //  A report type enumerator from hidpi.h,
               //  A pointer to a buffer for the returned array,
               //  The NumberInputValueCaps member of the device's HidP_Caps structure,
               //  A pointer to the PreparsedData structure returned by HidD_GetPreparsedData.

               //  Returns: True on success, False on failure.
               //  ***

               Int32 vcSize = Capabilities.NumberInputValueCaps;
               Byte[] valueCaps = new Byte[vcSize];

               NativeMethods.HidP_GetValueCaps(NativeMethods.HidP_Input, valueCaps, ref vcSize, preparsedData);

               // (To use this data, copy the ValueCaps byte array into an array of structures.)
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }
         finally
         {
            //  ***
            //  API function: HidD_FreePreparsedData

            //  Purpose: frees the buffer reserved by HidD_GetPreparsedData.

            //  Accepts: A pointer to the PreparsedData structure returned by HidD_GetPreparsedData.

            //  Returns: True on success, False on failure.
            //  ***

            if (preparsedData != IntPtr.Zero)
            {
               NativeMethods.HidD_FreePreparsedData(preparsedData);
            }
         }

         return Capabilities;
      }

      ///  <summary>
      ///  Creates a 32-bit Usage from the Usage Page and Usage ID.
      ///  Determines whether the Usage is a system mouse or keyboard.
      ///  Can be modified to detect other Usages.
      ///  </summary>
      ///
      ///  <param name="myCapabilities">A HIDP_CAPS structure retrieved with HidP_GetCaps.</param>
      ///
      ///  <returns>A String describing the Usage.</returns>

      internal String GetHidUsage(NativeMethods.HIDP_CAPS myCapabilities)
      {
         String usageDescription = "";

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "GetHidUsage");

         try
         {
            //  Create32-bit Usage from Usage Page and Usage ID.

            Int32 usage = myCapabilities.UsagePage * 256 + myCapabilities.Usage;

            if (usage == Convert.ToInt32(0X102))
            {
               usageDescription = "mouse";
            }
            if (usage == Convert.ToInt32(0X106))
            {
               usageDescription = "keyboard";
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return usageDescription;
      }

      ///  <summary>
      ///  Retrieves the number of Input reports the HID driver will store.
      ///  </summary>
      ///
      ///  <param name="hidHandle">A handle to a device.</param>
      ///  <param name="numberOfInputBuffers">An integer to hold the returned value.</param>
      ///
      ///  <returns>True on success. False on failure.</returns>

      internal Boolean GetNumberOfInputBuffers(SafeFileHandle hidHandle, ref Int32 numberOfInputBuffers)
      {
         Boolean success = false;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "GetNumberOfInputBuffers");

         try
         {
            //  ***
            //  API function: HidD_GetNumInputBuffers

            //  Purpose: retrieves the number of Input reports the host can store.
            //  Not supported by Windows 98 Gold.
            //  If the buffer is full and another report arrives, the host drops the
            //  ldest report.

            //  Accepts: a handle to a device and an integer to hold the number of buffers.

            //  Returns: True on success, False on failure.
            //  ***

            success = NativeMethods.HidD_GetNumInputBuffers(hidHandle, ref numberOfInputBuffers);
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return success;
      }

      ///  <summary>
      ///  Sets the number of Input reports the host HID driver store.
      ///  </summary>
      ///
      ///  <param name="hidHandle">A handle to the device.</param>
      ///  <param name="numberBuffers">The requested number of Input reports.</param>
      ///
      ///  <returns>True on success. False on failure.</returns>

      internal Boolean SetNumberOfInputBuffers(SafeFileHandle hidHandle, Int32 numberBuffers)
      {
         Boolean success = false;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "SetNumberOfInputBuffers");

         try
         {
            //  ***
            //  API function: HidD_SetNumInputBuffers

            //  Purpose: Sets the number of Input reports the host can store.
            //  If the buffer is full and another report arrives, the host drops the
            //  oldest report.

            //  Requires:
            //  A handle to a HID
            //  An integer to hold the number of buffers.

            //  Returns: true on success, false on failure.
            //  ***

            success = NativeMethods.HidD_SetNumInputBuffers(hidHandle, numberBuffers);
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return success;
      }

      ///  <summary>
      ///  Reads a Feature report from the device.
      ///  </summary>
      ///
      ///  <param name="hidHandle">The handle for learning about the device and exchanging Feature reports.</param>
      ///  <param name="inFeatureReportBuffer">Contains the requested report.</param>
      ///
      ///  <returns>True on success. False on failure.</returns>

      internal Boolean GetFeatureReport(SafeFileHandle hidHandle, ref Byte[] inFeatureReportBuffer)
      {
         Boolean success = false;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "GetFeatureReport");

         try
         {
            //  ***
            //  API function: HidD_GetFeature
            //  Attempts to read a Feature report from the device.

            //  Requires:
            //  A handle to a HID
            //  A pointer to a buffer containing the report ID and report
            //  The size of the buffer.

            //  Returns: true on success, false on failure.
            //  ***

            if (!hidHandle.IsInvalid && !hidHandle.IsClosed)
            {
               success = NativeMethods.HidD_GetFeature(hidHandle, inFeatureReportBuffer, inFeatureReportBuffer.Length);

               traceSource.TraceEvent(TraceEventType.Verbose, 1, " HidD_GetFeature success = {0}", success);
               traceSource.TraceEvent(TraceEventType.Information, 1, " Feature Report received: {0}", BitConverter.ToString(inFeatureReportBuffer));
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return success;
      }

      ///  <summary>
      ///  Writes a Feature report to the device.
      ///  </summary>
      ///
      ///  <param name="hidHandle">Handle to the device.</param>
      ///  <param name="outFeatureReportBuffer">Contains the report ID and report data.</param>
      ///
      ///  <returns>True on success. False on failure.</returns>

      internal Boolean SendFeatureReport(SafeFileHandle hidHandle, Byte[] outFeatureReportBuffer)
      {
         Boolean success = false;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "SendFeatureReport");

         try
         {
            //  ***
            //  API function: HidD_SetFeature

            //  Purpose: Attempts to send a Feature report to the device.

            //  Accepts:
            //  A handle to a HID
            //  A pointer to a buffer containing the report ID and report
            //  The size of the buffer.

            //  Returns: true on success, false on failure.
            //  ***

            if (!hidHandle.IsInvalid && !hidHandle.IsClosed)
            {
               success = NativeMethods.HidD_SetFeature(hidHandle, outFeatureReportBuffer, outFeatureReportBuffer.Length);

               traceSource.TraceEvent(TraceEventType.Verbose, 1, " HidD_SetFeature success = {0}", success);
               traceSource.TraceEvent(TraceEventType.Information, 1, " Feature Report sent: {0}", BitConverter.ToString(outFeatureReportBuffer));
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return success;
      }

      ///  <summary>
      ///  Reads an Input report from the device using a control transfer.
      ///  </summary>
      ///
      ///  <param name="hidHandle">The handle for learning about the device and exchanging Feature reports.</param>
      ///  <param name="inputReportBuffer">Contains the requested report.</param>
      ///
      ///  <returns>True on success. False on failure.</returns>

      internal Boolean GetInputReportViaControlTransfer(SafeFileHandle hidHandle, ref Byte[] inputReportBuffer)
      {
         Boolean success = false;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "GetInputReportViaControlTransfer");

         try
         {
            //  ***
            //  API function: HidD_GetInputReport

            //  Purpose: Attempts to read an Input report from the device using a control transfer.

            //  Requires:
            //  A handle to a HID
            //  A pointer to a buffer containing the report ID and report
            //  The size of the buffer.

            //  Returns: true on success, false on failure.
            //  ***

            if (!hidHandle.IsInvalid && !hidHandle.IsClosed)
            {
               success = NativeMethods.HidD_GetInputReport(hidHandle, inputReportBuffer, inputReportBuffer.Length);

               traceSource.TraceEvent(TraceEventType.Verbose, 1, " HidD_GetInputReport success = {0}", success);
               traceSource.TraceEvent(TraceEventType.Information, 1, " Input Report received: {0}", BitConverter.ToString(inputReportBuffer));
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return success;
      }

      ///  <summary>
      ///  Writes an Output report to the device using a control transfer.
      ///  </summary>
      ///
      ///  <param name="hidHandle">Handle to the device.</param>
      ///  <param name="outputReportBuffer">Contains the report ID and report data.</param>
      ///
      ///  <returns>True on success. False on failure.</returns>

      internal Boolean SendOutputReportViaControlTransfer(SafeFileHandle hidHandle, Byte[] outputReportBuffer)
      {
         Boolean success = false;

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "SendOutputReportViaControlTransfer");

         try
         {
            //  ***
            //  API function: HidD_SetOutputReport

            //  Purpose:
            //  Attempts to send an Output report to the device using a control transfer.

            //  Accepts:
            //  A handle to a HID
            //  A pointer to a buffer containing the report ID and report
            //  The size of the buffer.

            //  Returns: true on success, false on failure.
            //  ***

            if (!hidHandle.IsInvalid && !hidHandle.IsClosed)
            {
               success = NativeMethods.HidD_SetOutputReport(hidHandle, outputReportBuffer, outputReportBuffer.Length);

               traceSource.TraceEvent(TraceEventType.Verbose, 1, " HidD_SetOutputReport success = {0}", success);
               traceSource.TraceEvent(TraceEventType.Information, 1, " Output Report sent: {0}", BitConverter.ToString(outputReportBuffer));
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return success;
      }

      ///  <summary>
      ///  Get the manufacturer string of the device.
      ///  </summary>
      ///
      ///  <param name="hidHandle">A handle for accessing the device.</param>
      ///
      ///  <returns>Manufacturer string on success. An empty string on failure.</returns>

      internal String GetManufacturerString(SafeFileHandle hidHandle)
      {
         Boolean success = false;
         Byte[] buffer = new Byte[254];
         String manufacturerString = "";

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "GetManufacturerString");

         try
         {
            //  ***
            //  API function: HidD_GetManufacturerString

            //  Purpose: Attempts to read the top-level collection's embedded string that identifies the manufacturer.

            //  Requires:
            //  A handle to a HID
            //  A pointer to a buffer to store the string
            //  The size of the buffer (the maximum length is 127 for USB devices, including NULL at the end).

            //  Returns: true on success, false on failure.
            //  ***

            if (!hidHandle.IsInvalid && !hidHandle.IsClosed)
            {
               success = NativeMethods.HidD_GetManufacturerString(hidHandle, buffer, buffer.Length);
               
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " HidD_GetManufacturerString success = {0}", success);
               
               if(success)
               {
                  manufacturerString = System.Text.Encoding.Unicode.GetString(buffer);
                  
                  traceSource.TraceEvent(TraceEventType.Information, 1, " Manufacturer string received: {0}", manufacturerString);
               }
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return manufacturerString.Trim(new [] {'\0'});
      }

      ///  <summary>
      ///  Get the product string of the device.
      ///  </summary>
      ///
      ///  <param name="hidHandle">A handle for accessing the device.</param>
      ///
      ///  <returns>Product string on success. An empty string on failure.</returns>

      internal String GetProductString(SafeFileHandle hidHandle)
      {
         Boolean success = false;
         Byte[] buffer = new Byte[254];
         String productString = "";

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "GetProductString");

         try
         {
            //  ***
            //  API function: HidD_GetProductString

            //  Purpose: Attempts to read the top-level collection's embedded string that identifies the product.

            //  Requires:
            //  A handle to a HID
            //  A pointer to a buffer to store the string
            //  The size of the buffer (the maximum length is 127 for USB devices, including NULL at the end).

            //  Returns: true on success, false on failure.
            //  ***

            if (!hidHandle.IsInvalid && !hidHandle.IsClosed)
            {
               success = NativeMethods.HidD_GetProductString(hidHandle, buffer, buffer.Length);
               
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " HidD_GetProductString success = {0}", success);
               
               if(success)
               {
                  productString = System.Text.Encoding.Unicode.GetString(buffer);
                  
                  traceSource.TraceEvent(TraceEventType.Information, 1, " Product string received: {0}", productString);
               }
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return productString.Trim(new [] {'\0'});
      }

      ///  <summary>
      ///  Get the serial number string of the device.
      ///  </summary>
      ///
      ///  <param name="hidHandle">A handle for accessing the device.</param>
      ///
      ///  <returns>Serial number string on success. An empty string on failure.</returns>

      internal String GetSerialNumberString(SafeFileHandle hidHandle)
      {
         Boolean success = false;
         Byte[] buffer = new Byte[254];
         String serialNumberString = "";

         traceSource.TraceEvent(TraceEventType.Verbose, 1, "GetSerialNumberString");

         try
         {
            //  ***
            //  API function: HidD_GetSerialNumberString

            //  Purpose: Attempts to read the top-level collection's embedded string that identifies the serial number.

            //  Requires:
            //  A handle to a HID
            //  A pointer to a buffer to store the string
            //  The size of the buffer (the maximum length is 127 for USB devices, including NULL at the end).

            //  Returns: true on success, false on failure.
            //  ***

            if (!hidHandle.IsInvalid && !hidHandle.IsClosed)
            {
               success = NativeMethods.HidD_GetSerialNumberString(hidHandle, buffer, buffer.Length);
               
               traceSource.TraceEvent(TraceEventType.Verbose, 1, " HidD_GetSerialNumberString success = {0}", success);
               
               if(success)
               {
                  serialNumberString = System.Text.Encoding.Unicode.GetString(buffer);
                  
                  traceSource.TraceEvent(TraceEventType.Information, 1, " Serial number string received: {0}", serialNumberString);
               }
            }
         }
         catch (Exception ex)
         {
            DisplayException(MethodBase.GetCurrentMethod().Name, ex, ShowMsgBoxOnException);
         }

         return serialNumberString.Trim(new [] {'\0'});
      }
   }
}
