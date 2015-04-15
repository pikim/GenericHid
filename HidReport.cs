using System;

namespace GenericHid
{
    public class HidReport
    {
        private byte _reportId;
        private byte[] _data;

        public HidReport(int reportSize)
        {
          _data = new byte[reportSize-1];
        }

        public HidReport(byte reportId, byte reportData)
        {
          _reportId = reportId;
          _data = new byte[1];
          _data[0] = reportData;
        }

        public HidReport(byte reportId, byte[] reportData)
        {
          if (reportData != null)
          {
            _reportId = reportId;
            _data = new byte[reportData.Length];
            Buffer.BlockCopy(reportData, 0, _data, 0, reportData.Length);
          }
        }

//        public HidReport(int reportSize, byte reportId, byte[] reportData)
//        {
//          if (reportData != null)
//          {
//            _reportId = reportId;
//            _data = new byte[reportData.Length];
//            Buffer.BlockCopy(reportData, 0, _data, 0, Math.Min(reportSize, reportData.Length));
//          }
//        }

        public HidReport(byte[] report)
        {
          if (report != null)
          {
            _reportId = report[0];
            _data = new byte[report.Length-1];
            Buffer.BlockCopy(report, 1, _data, 0, report.Length - 1);
          }
        }

        public byte ReportId
        {
          get { return _reportId; }
          set { _reportId = value; }
        }

        public byte[] Data
        {
          get { return _data; }
          set { _data = value; }
        }

        public short Length
        {
          get { return Convert.ToInt16(_data.Length+1); }
        }

        public byte[] GetBytes()
        {
          byte[] data = new byte[_data.Length+1];
          data[0] = _reportId;
          Buffer.BlockCopy(_data, 0, data, 1, _data.Length);
          return data;
        }

        public byte[] GetBytes(int reportByteLength)
        {
          int _length = Math.Min(reportByteLength, _data.Length+1);
          byte[] data = new byte[_length];
          data[0] = _reportId;
          Buffer.BlockCopy(_data, 0, data, 1, _length-1);
          return data;
        }
    }
}
