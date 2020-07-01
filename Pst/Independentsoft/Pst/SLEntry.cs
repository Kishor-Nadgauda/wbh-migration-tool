using System;

namespace Independentsoft.Pst
{
    internal class SLEntry
    {
        private ulong nid;
        private ulong bidData;
        private ulong bidSub;
        
        internal SLEntry()
        {
        }

        internal SLEntry(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            if (buffer.Length == 24)
            {
                nid = Util.ToUInt64(buffer, 0);
                bidData = Util.ToUInt64(buffer, 8);
                bidSub = Util.ToUInt64(buffer, 16);
            }
            else
            {
                nid = BitConverter.ToUInt32(buffer, 0);
                bidData = BitConverter.ToUInt32(buffer, 4);
                bidSub = BitConverter.ToUInt32(buffer, 8);
            }
        }

        #region Propertis

        internal ulong NodeId
        {
            get
            {
                return nid;
            }
        }

        internal ulong BlockId
        {
            get
            {
                return bidData;
            }
        }

        internal ulong SubNodeBlockId
        {
            get
            {
                return bidSub;
            }
        }

        #endregion
    }
}
