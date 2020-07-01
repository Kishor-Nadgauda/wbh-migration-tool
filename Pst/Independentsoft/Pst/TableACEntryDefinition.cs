using System;

namespace Independentsoft.Pst
{
    internal class TableACEntryDefinition
    {  
        private ushort entryValueType;
        private ushort entryId;
        private ushort valueArrayEntryOffset;
        private ushort valueArrayEntrySize;
        private ushort valueArrayEntryNumber;
        private uint descriptorId;

        internal TableACEntryDefinition()
        {
        }

        #region Properties

        internal ushort EntryValueType
        {
            get
            {
                return entryValueType;
            }
            set
            {
                entryValueType = value;
            }
        }

        internal ushort EntryId
        {
            get
            {
                return entryId;
            }
            set
            {
                entryId = value;
            }
        }

        internal ushort ValueArrayEntryOffset
        {
            get
            {
                return valueArrayEntryOffset;
            }
            set
            {
                valueArrayEntryOffset = value;
            }
        }

        internal ushort ValueArrayEntrySize
        {
            get
            {
                return valueArrayEntrySize;
            }
            set
            {
                valueArrayEntrySize = value;
            }
        }

        internal ushort ValueArrayEntryNumber
        {
            get
            {
                return valueArrayEntryNumber;
            }
            set
            {
                valueArrayEntryNumber = value;
            }
        }

        internal uint DescriptorId
        {
            get
            {
                return descriptorId;
            }
            set
            {
                descriptorId = value;
            }
        }

        #endregion
    }
}
