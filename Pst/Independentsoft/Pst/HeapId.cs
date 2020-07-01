using System;

namespace Independentsoft.Pst
{
    internal class HeapId
    {
        private byte hidType;
        private short hidIndex;
        private ushort hidBlockIndex;

        internal HeapId()
        {
        }

        internal HeapId(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            ushort hid = BitConverter.ToUInt16(buffer, 0);

            hidType = (byte)(hid << 11);
            hidIndex = (short)(hid >> 5);

            hidBlockIndex = BitConverter.ToUInt16(buffer, 2);
        }

        #region Properties

        internal byte Type
        {
            get
            {
                return hidType;
            }
        }

        internal short Index
        {
            get
            {
                return hidIndex;
            }
        }

        internal ushort BlockIndex
        {
            get
            {
                return hidBlockIndex;
            }
        }

        #endregion
    }
}
