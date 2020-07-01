using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    internal class TableHeader7C
    {  
        private byte type;
        private byte entryCount;
        private ushort valueArraySize;
        private uint b5HeaderReference;
        private uint valueEntriesIndexReference;
        private IList<Table7CEntryDefinition> entryDefinitions = new List<Table7CEntryDefinition>();

        internal TableHeader7C()
        {
        }

        internal TableHeader7C(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            if (buffer.Length > 0)
            {
                this.type = buffer[0];

                if (type != 0x7C)
                {
                    return;
                }

                this.entryCount = buffer[1];
                this.valueArraySize = BitConverter.ToUInt16(buffer, 8);
                this.b5HeaderReference = BitConverter.ToUInt32(buffer, 10);
                this.valueEntriesIndexReference = BitConverter.ToUInt32(buffer, 14);

                for (int i = 0; i < entryCount; i++)
                {
                    Table7CEntryDefinition entry = new Table7CEntryDefinition();

                    entry.EntryValueType = BitConverter.ToUInt16(buffer, 22 + i * 8);
                    entry.EntryId = BitConverter.ToUInt16(buffer, 22 + i * 8 + 2);
                    entry.ValueArrayEntryOffset = BitConverter.ToUInt16(buffer, 22 + i * 8 + 4);
                    entry.ValueArrayEntrySize = buffer[22 + i * 8 + 6];
                    entry.ValueArrayEntryNumber = buffer[22 + i * 8 + 7];

                    entryDefinitions.Add(entry);
                }
            }
        }

        #region Properties

        internal byte Type
        {
            get
            {
                return type;
            }
        }

        internal byte EntryCount
        {
            get
            {
                return entryCount;
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
            set
            {
                valueEntriesIndexReference = value;
            }

        }

        internal IList<Table7CEntryDefinition> EntryDefinitions
        {
            get
            {
                return entryDefinitions;
            }
        }

        #endregion
    }
}
