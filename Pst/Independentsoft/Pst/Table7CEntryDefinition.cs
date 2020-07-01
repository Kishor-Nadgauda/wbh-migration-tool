using System;

namespace Independentsoft.Pst
{
    internal class Table7CEntryDefinition
    {  
        private ushort entryValueType;
        private ushort entryId;
        private ushort valueArrayEntryOffset;
        private byte valueArrayEntrySize;
        private byte valueArrayEntryNumber;

        internal Table7CEntryDefinition()
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

        internal byte ValueArrayEntrySize
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

        internal byte ValueArrayEntryNumber
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

        #endregion
    }
}
