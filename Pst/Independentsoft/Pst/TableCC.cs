using System;
using System.IO;

namespace Independentsoft.Pst
{
    internal class TableCC : Table
    {
        private TableIndex index;
        private byte[] tableBuffer;

        internal TableCC()
        {
        }

        internal TableCC(PstFileReader reader, byte[] tableBuffer)
        {
            this.tableBuffer = tableBuffer;

            Parse(reader, tableBuffer);
        }

        private void Parse(PstFileReader reader, byte[] tableBuffer)
        {
            ushort indexOffset = BitConverter.ToUInt16(tableBuffer, 0);
            ushort tableType = BitConverter.ToUInt16(tableBuffer, 2);
            uint tableValueReference = BitConverter.ToUInt32(tableBuffer, 4);

            if (tableType != 0xCC)
            {
                return;
            }

            index = new TableIndex(tableBuffer, indexOffset, reader.PstFile.Is64Bit); 
        }

        internal byte[] GetValue(int itemIndex)
        {
            if (index == null || itemIndex > index.Items.Count)
            {
                return null;
            }
            else
            {
                TableIndexItem indexItem = index.Items.ContainsKey(itemIndex) ? index.Items[itemIndex] : null;

                byte[] valueBuffer = new byte[indexItem.EndOffset - indexItem.StartOffset];
                System.Array.Copy(this.tableBuffer, indexItem.StartOffset, valueBuffer, 0, valueBuffer.Length);

                return valueBuffer;
            }
        }
    }
}
