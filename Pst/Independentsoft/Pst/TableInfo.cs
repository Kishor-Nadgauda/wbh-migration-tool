using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    internal class TableInfo
    {
        private byte bType;
        private byte cCols;
        private HeapId hidRowIndex;
        private uint hnidRows;
        private HeapId hidIndex;
        private IList<TableColumnDescription> columns = new List<TableColumnDescription>();

        private ushort endOffset4ByteDataValue;
        private ushort endOffset2ByteDataValue;
        private ushort endOffset1ByteDataValue;
        private ushort endOffsetCellExistenceBlock;

        internal TableInfo()
        {
        }

        internal TableInfo(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            bType = buffer[0];
            cCols = buffer[1];

            endOffset4ByteDataValue = BitConverter.ToUInt16(buffer, 2);
            endOffset2ByteDataValue = BitConverter.ToUInt16(buffer, 4);
            endOffset1ByteDataValue = BitConverter.ToUInt16(buffer, 6);
            endOffsetCellExistenceBlock = BitConverter.ToUInt16(buffer, 8);

            byte[] hidRowIndexBuffer = new byte[4];
            System.Array.Copy(buffer, 10, hidRowIndexBuffer, 0, 4);

            hidRowIndex = new HeapId(hidRowIndexBuffer);

            hnidRows = BitConverter.ToUInt32(buffer, 14);

            byte[] hidIndexBuffer = new byte[4];
            System.Array.Copy(buffer, 18, hidIndexBuffer, 0, 4);

            hidIndex = new HeapId(hidIndexBuffer);

            for (int i = 0; i < cCols; i++)
            {
                byte[] tableColumnDescriptionBuffer = new byte[8];
                System.Array.Copy(buffer, 22 + 8*i, tableColumnDescriptionBuffer, 0, 8);

                TableColumnDescription tableColumnDescription = new TableColumnDescription(tableColumnDescriptionBuffer);
                columns.Add(tableColumnDescription);
            }
        }

        #region Properties

        internal HeapId HidRowIndex
        {
            get
            {
                return hidRowIndex;
            }
        }

        internal uint HnidRows
        {
            get
            {
                return hnidRows;
            }
        }

        internal IList<TableColumnDescription> Columns
        {
            get
            {
                return columns;
            }
        }

        internal ushort EndOffset4ByteDataValue
        {
            get
            {
                return endOffset4ByteDataValue;
            }
        }

        internal ushort EndOffset2ByteDataValue
        {
            get
            {
                return endOffset2ByteDataValue;
            }
        }

        internal ushort EndOffset1ByteDataValue
        {
            get
            {
                return endOffset1ByteDataValue;
            }
        }

        internal ushort EndOffsetCellExistenceBlock
        {
            get
            {
                return endOffsetCellExistenceBlock;
            }
        }

        #endregion
    }
}
