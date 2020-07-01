using System;
using System.IO;

namespace Independentsoft.Pst
{
    /// <summary>
    /// The table contains GUID descriptors.
    /// </summary>
    internal class Table9C : Table
    {
        internal Table9C()
        {
        }

        internal Table9C(PstFileReader reader, byte[] tableBuffer, LocalDescriptorList localDescriptorList, DataStructure tableDataNode)
        {
            Parse(reader, tableBuffer, localDescriptorList, tableDataNode);
        }

        private void Parse(PstFileReader reader, byte[] tableBuffer, LocalDescriptorList localDescriptorList, DataStructure tableDataNode)
        {
            ushort indexOffset = BitConverter.ToUInt16(tableBuffer, 0);
            ushort tableType = BitConverter.ToUInt16(tableBuffer, 2);
            uint tableValueReference = BitConverter.ToUInt32(tableBuffer, 4);

            TableIndex index = new TableIndex(tableBuffer, indexOffset, reader.PstFile.Is64Bit);

            TableIndexItem tableHeader9CIndexItem = index.GetItem(tableValueReference);

            uint tableHeaderB5Reference = BitConverter.ToUInt32(tableBuffer, tableHeader9CIndexItem.StartOffset);

            TableIndexItem tableHeaderB5IndexItem = index.GetItem(tableHeaderB5Reference);

            byte[] tableHeaderB5Buffer = new byte[tableHeaderB5IndexItem.EndOffset - tableHeaderB5IndexItem.StartOffset];
            System.Array.Copy(tableBuffer, tableHeaderB5IndexItem.StartOffset, tableHeaderB5Buffer, 0, tableHeaderB5Buffer.Length);

            TableHeaderB5 tableHeaderB5 = new TableHeaderB5(tableHeaderB5Buffer);

            if (tableHeaderB5.Type != 0xB5)
            {
                return;
            }

            if (tableHeaderB5.ValueEntriesIndexReference == 0)
            {
                return;
            }

            TableIndexItem valueEntiresIndexItem = index.GetItem(tableHeaderB5.ValueEntriesIndexReference);

            for (int i = 0; i < valueEntiresIndexItem.EndOffset - valueEntiresIndexItem.StartOffset; i += 20)
            {
                int startIndex = valueEntiresIndexItem.StartOffset + i;

                byte[] guid = new byte[16];
                System.Array.Copy(tableBuffer, startIndex, guid, 0, 16);

                uint descriptorId = BitConverter.ToUInt16(tableBuffer, startIndex + 16);
                
            }
        }
    }
}
