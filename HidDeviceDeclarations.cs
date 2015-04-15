using System;

namespace GenericHid
{
   public sealed partial class HidDevice
   {
      public delegate void DeviceNotifyDelegate();
      public delegate void DataReceiveDelegate(HidReport report);

      private enum ReportAction
      {
         Read,
         Written
      }

      private enum ReportTypes
      {
         Input,
         Output,
         Feature
      }

      public enum TransferTypes
      {
         Control,
         Interrupt
      }

      private enum WmiDeviceProperties
      {
         Name,
         Caption,
         Description,
         Manufacturer,
         PNPDeviceID,
         DeviceID,
         ClassGUID
      }
   }
}
