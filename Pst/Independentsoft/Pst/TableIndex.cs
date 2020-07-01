using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    internal class TableIndex
    {
        private byte[] tableBuffer;
        private ushort indexOffset;
        private ushort itemCount;
        private IDictionary<int, TableIndexItem> items;

        uint[] secondIndexOffset = new uint[0];
        private IDictionary<int, TableIndexItem>[] secondItems = new Dictionary<int, TableIndexItem>[0];
        private bool is64Bit;

        internal TableIndex()
        {
        }

        internal TableIndex(byte[] tableBuffer, ushort indexOffset, bool is64Bit)
        {
            this.is64Bit = is64Bit;

            this.tableBuffer = tableBuffer;
            this.indexOffset = indexOffset;

            if (indexOffset >= tableBuffer.Length)
            {
              this.itemCount = 0;
            }
            else
            {
              this.itemCount = (ushort)(BitConverter.ToUInt16(tableBuffer, (int)indexOffset) + 1);
            }

            if (this.itemCount > 0)
            {
                items = new Dictionary<int, TableIndexItem>();
            }

            for (int i = 0; i < itemCount; i++)
            {
                int offset = (int)(indexOffset + 2) + i * 2;

                TableIndexItem item = new TableIndexItem();

                if (tableBuffer.Length > offset + 3)
                {
                    ushort startOffset = BitConverter.ToUInt16(tableBuffer, offset);
                    ushort endOffset = BitConverter.ToUInt16(tableBuffer, offset + 2);

                    item.StartOffset = startOffset;
                    item.EndOffset = endOffset;

                    items.Add(i, item);
                }
            }

            int tableSize = is64Bit ? 8176: 8180;

            //workaround
            if (tableBuffer.Length > tableSize)
            {
                int secondTableCount = (int)Math.Ceiling((double)(tableBuffer.Length - tableSize) / tableSize);

                secondItems = new Dictionary<int, TableIndexItem>[secondTableCount];
                secondIndexOffset = new uint[secondTableCount];

                for (int s = 0; s < secondTableCount; s++)
                {
                    secondItems[s] = new Dictionary<int, TableIndexItem>();

                    int currentTableOffset = tableSize * (s + 1);

                    byte[] currentTableBuffer = s < secondTableCount - 1 ? new byte[tableSize] : new byte[tableBuffer.Length - currentTableOffset];
    
                    System.Array.Copy(tableBuffer, currentTableOffset, currentTableBuffer, 0, currentTableBuffer.Length);

                    uint currentIndexOffset = BitConverter.ToUInt16(currentTableBuffer, 0);
                    secondIndexOffset[s] = (uint)(currentTableOffset + currentIndexOffset);

                    ushort currentItemCount = (ushort)(BitConverter.ToUInt16(currentTableBuffer, (int)currentIndexOffset) + 1);

                    for (int i = 0; i < currentItemCount; i++)
                    {
                        int offset = (int)(currentIndexOffset + 2) + i * 2;

                        TableIndexItem item = new TableIndexItem();

                        ushort startOffset = BitConverter.ToUInt16(currentTableBuffer, offset);
                        ushort endOffset = BitConverter.ToUInt16(currentTableBuffer, offset + 2);

                        item.StartOffset = startOffset + currentTableOffset;
                        item.EndOffset = endOffset + currentTableOffset;

                        if (!items.ContainsKey(itemCount))
                        {
                            items.Add(itemCount, item);
                            itemCount++;
                        }

                        if (!secondItems[s].ContainsKey(i))
                        {
                            secondItems[s].Add(i, item);
                        }
                    }
                }
            }
        }

        internal TableIndexItem GetItem(uint reference)
        {
            TableIndexItem item = new TableIndexItem();

            int itemOffset = (int)(indexOffset + 2 + (reference >> 4));
            int tableIndex = (int)Math.Floor((double)(reference / 65536));

            if (reference > 65536 && secondIndexOffset != null && secondIndexOffset.Length >= tableIndex)
            {
                itemOffset = (int)(secondIndexOffset[tableIndex - 1] + 2 + ((reference - 65536 * tableIndex) >> 4));
            }

            if (itemOffset > tableBuffer.Length - 4)
            {
                return null;
            }

            int startOffset = BitConverter.ToUInt16(tableBuffer, itemOffset);
            int endOffset = BitConverter.ToUInt16(tableBuffer, itemOffset + 2);

            int tableSize = is64Bit ? 8176 : 8180;
            int secondTableOffset = tableIndex == 0 ? tableSize : tableIndex * tableSize;

            if (itemOffset > secondTableOffset)
            {
                startOffset = (int)(startOffset + secondTableOffset);
                endOffset = (int)(endOffset + secondTableOffset);
            }

            if (endOffset < startOffset)
            {
                return null;
            }
            else
            {
                item.StartOffset = startOffset;
                item.EndOffset = endOffset;

                return item;
            }
        }

        #region Properties

        internal ushort ItemCount
        {
            get
            {
                return itemCount;
            }
        }

        internal IDictionary<int, TableIndexItem> Items
        {
            get
            {
                return items;
            }
        }

        internal IDictionary<int, TableIndexItem>[] SecondItems
        {
            get
            {
                return secondItems;
            }
        }

        #endregion
    }
}
