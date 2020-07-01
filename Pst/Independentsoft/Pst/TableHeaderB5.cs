using System;

namespace Independentsoft.Pst
{
    internal class TableHeaderB5
    {  
        private byte type;
        private byte entryIdSize;
        private byte entryValueSize;
        private uint valueEntriesIndexReference;

        internal TableHeaderB5()
        {
        }

        internal TableHeaderB5(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            this.type = buffer[0];
            this.entryIdSize = buffer[1];
            this.entryValueSize = buffer[2];
            this.valueEntriesIndexReference = BitConverter.ToUInt32(buffer, 4);
        }

        #region Properties

        internal byte Type
        {
            get
            {
                return type;
            }
        }

        internal byte EntryIdSize
        {
            get
            {
                return entryIdSize;
            }
        }

        internal byte EntryValueSize
        {
            get
            {
                return entryValueSize;
            }
        }

        internal uint ValueEntriesIndexReference
        {
            get
            {
                return valueEntriesIndexReference;
            }
        }

        #endregion
    }
}
