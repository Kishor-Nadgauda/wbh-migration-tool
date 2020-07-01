using System;

namespace Independentsoft.Pst
{
    internal class TableColumnDescription
    {
        private uint tag;
        private ushort ibData;
        private byte cbData;
        private byte iBit;

        internal TableColumnDescription()
        {
        }

        internal TableColumnDescription(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            tag = BitConverter.ToUInt32(buffer, 0);
            ibData = BitConverter.ToUInt16(buffer, 4);
            cbData = buffer[6];
            iBit = buffer[7];
        }

        #region Properties

        internal uint Tag
        {
            get
            {
                return tag;
            }
        }

        internal ushort Offset
        {
            get
            {
                return ibData;
            }
        }

        internal byte CountBytes
        {
            get
            {
                return cbData;
            }
        }

        internal byte IBit
        {
            get
            {
                return iBit;
            }
        }

        #endregion
    }
}
