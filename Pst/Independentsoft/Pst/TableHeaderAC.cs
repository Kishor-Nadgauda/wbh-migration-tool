using System;

namespace Independentsoft.Pst
{
    internal class TableHeaderAC
    {  
        private byte type;
        private ushort valueArraySize;
        private uint b5HeaderReference;
        private uint valueEntriesIndexReference;
        private ushort entryDefinitionCount;
        private uint tableEntryDefinitionReference;

        internal TableHeaderAC()
        {
        }

        internal TableHeaderAC(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            this.type = buffer[0];
            this.valueArraySize = BitConverter.ToUInt16(buffer, 8);
            this.b5HeaderReference = BitConverter.ToUInt32(buffer, 10);
            this.valueEntriesIndexReference = BitConverter.ToUInt32(buffer, 14);
            this.entryDefinitionCount = BitConverter.ToUInt16(buffer, 22);
            this.tableEntryDefinitionReference = BitConverter.ToUInt32(buffer, 24);
        }

        #region Properties

        internal byte Type
        {
            get
            {
                return type;
            }
        }

        internal ushort ValueArraySize
        {
            get
            {
                return valueArraySize;
            }
        }

        internal uint B5HeaderReference
        {
            get
            {
                return b5HeaderReference;
            }
        }

        internal uint ValueEntriesIndexReference
        {
            get
            {
                return valueEntriesIndexReference;
            }
        }

        internal ushort EntryDefinitionCount
        {
            get
            {
                return entryDefinitionCount;
            }
        }

        internal uint TableEntryDefinitionReference
        {
            get
            {
                return tableEntryDefinitionReference;
            }
        }

        #endregion
    }
}
