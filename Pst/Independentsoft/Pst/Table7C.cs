using System;
using System.IO;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    internal class Table7C : Table
    {
        internal Table7C()
        {
        }

        internal Table7C(PstFileReader reader, byte[] tableBuffer, LocalDescriptorList localDescriptorList, DataStructure tableDataNode)
        {
            Parse(reader, tableBuffer, localDescriptorList, tableDataNode, false);
        }

        internal Table7C(PstFileReader reader, byte[] tableBuffer, LocalDescriptorList localDescriptorList, DataStructure tableDataNode, bool isRecipientsTable)
        {
            Parse(reader, tableBuffer, localDescriptorList, tableDataNode, isRecipientsTable);
        }

        private void Parse(PstFileReader reader, byte[] tableBuffer, LocalDescriptorList localDescriptorList, DataStructure tableDataNode, bool isRecipientsTable)
        {
            ushort indexOffset = BitConverter.ToUInt16(tableBuffer, 0);
            ushort tableType = BitConverter.ToUInt16(tableBuffer, 2);
            uint tableValueReference = BitConverter.ToUInt32(tableBuffer, 4);

            TableIndex index = new TableIndex(tableBuffer, indexOffset, reader.PstFile.Is64Bit);
         
            TableIndexItem tableHeader7CIndexItem = index.GetItem(tableValueReference);

            if (tableHeader7CIndexItem == null)
            {
                return;
            }

            byte[] tableHeader7CBuffer = new byte[tableHeader7CIndexItem.EndOffset - tableHeader7CIndexItem.StartOffset];

            if (tableHeader7CBuffer.Length > tableBuffer.Length - tableHeader7CIndexItem.StartOffset)
            {
                return;
            }

            System.Array.Copy(tableBuffer, tableHeader7CIndexItem.StartOffset, tableHeader7CBuffer, 0, tableHeader7CBuffer.Length);

            TableHeader7C tableHeader7C = new TableHeader7C(tableHeader7CBuffer);

            if (tableHeader7C.Type != 0x7C)
            {
                return;
            }

            TableIndexItem tableHeaderB5IndexItem = index.GetItem(tableHeader7C.B5HeaderReference);

            byte[] tableHeaderB5Buffer = new byte[8];
            System.Array.Copy(tableBuffer, tableHeaderB5IndexItem.StartOffset, tableHeaderB5Buffer, 0, 8);

            TableHeaderB5 tableHeaderB5 = new TableHeaderB5(tableHeaderB5Buffer);
            BTreeOnHeapHeader bTreeOnHeapHeader = new BTreeOnHeapHeader(tableHeaderB5Buffer);

            if (tableHeaderB5.Type != 0xB5)
            {
                return;
            }

            if (tableHeader7C.ValueEntriesIndexReference == 0)
            {
                return;
            }

            byte[] entryValueBuffer = null;
            bool useB5Header = false;

            if (bTreeOnHeapHeader.BIdxLevels == 2)
            {
                //Not implemented
                return;
            }
            else if (bTreeOnHeapHeader.BIdxLevels == 1)
            {
                byte[] level1EntryValueBuffer = null;

                if (bTreeOnHeapHeader.HidRoot.Index > 0 && bTreeOnHeapHeader.HidRoot.BlockIndex == 0)
                {
                    TableIndexItem entryIndexItem = index.Items.ContainsKey((int)bTreeOnHeapHeader.HidRoot.Index) ? index.Items[(int)bTreeOnHeapHeader.HidRoot.Index] : null;

                    if (entryIndexItem != null)
                    {
                        level1EntryValueBuffer = new byte[entryIndexItem.EndOffset - entryIndexItem.StartOffset];
                        System.Array.Copy(tableBuffer, entryIndexItem.StartOffset, level1EntryValueBuffer, 0, level1EntryValueBuffer.Length);
                    }
                }
                else if (bTreeOnHeapHeader.HidRoot.Index > 0 && bTreeOnHeapHeader.HidRoot.BlockIndex > 0)
                {
                    if (index.SecondItems.Length >= bTreeOnHeapHeader.HidRoot.BlockIndex)
                    {
                        IDictionary<int, TableIndexItem> secondIndexTable = index.SecondItems[bTreeOnHeapHeader.HidRoot.BlockIndex - 1];

                        if (secondIndexTable != null)
                        {
                            TableIndexItem entryIndexItem = secondIndexTable.ContainsKey((int)bTreeOnHeapHeader.HidRoot.Index) ? secondIndexTable[(int)bTreeOnHeapHeader.HidRoot.Index] : null;

                            if (entryIndexItem != null)
                            {
                                level1EntryValueBuffer = new byte[entryIndexItem.EndOffset - entryIndexItem.StartOffset];
                                System.Array.Copy(tableBuffer, entryIndexItem.StartOffset, level1EntryValueBuffer, 0, level1EntryValueBuffer.Length);
                            }
                        }
                    }
                }

                MemoryStream entryValueBufferMemoryStream = new MemoryStream();

                for (int i = 0; i < level1EntryValueBuffer.Length; i+=8)
                {
                   byte[] heapIdBuffer = new byte[4];
                    System.Array.Copy(level1EntryValueBuffer, i+4, heapIdBuffer, 0, 4);

                    HeapId heapId = new HeapId(heapIdBuffer);

                    if (heapId.Index > 0 && heapId.BlockIndex == 0)
                    {
                        TableIndexItem tempEntryIndexItem = index.Items.ContainsKey((int)heapId.Index) ? index.Items[(int)heapId.Index] : null;

                        if (tempEntryIndexItem != null)
                        {
                            byte[] tempLevel1EntryValueBuffer = new byte[tempEntryIndexItem.EndOffset - tempEntryIndexItem.StartOffset];
                            System.Array.Copy(tableBuffer, tempEntryIndexItem.StartOffset, tempLevel1EntryValueBuffer, 0, tempLevel1EntryValueBuffer.Length);

                            entryValueBufferMemoryStream.Write(tempLevel1EntryValueBuffer, 0, tempLevel1EntryValueBuffer.Length);
                        }
                    }
                    else if (heapId.Index > 0 && heapId.BlockIndex > 0)
                    {
                        if (index.SecondItems.Length >= heapId.BlockIndex)
                        {
                            IDictionary<int, TableIndexItem> secondIndexTable = index.SecondItems[heapId.BlockIndex - 1];

                            if (secondIndexTable != null)
                            {
                                TableIndexItem tempEntryIndexItem = secondIndexTable.ContainsKey((int)heapId.Index) ? secondIndexTable[(int)heapId.Index] : null;

                                if (tempEntryIndexItem != null)
                                {
                                    byte[] tempLevel1EntryValueBuffer = new byte[tempEntryIndexItem.EndOffset - tempEntryIndexItem.StartOffset];
                                    System.Array.Copy(tableBuffer, tempEntryIndexItem.StartOffset, tempLevel1EntryValueBuffer, 0, tempLevel1EntryValueBuffer.Length);

                                    entryValueBufferMemoryStream.Write(tempLevel1EntryValueBuffer, 0, tempLevel1EntryValueBuffer.Length);
                                }
                            }
                        }
                    }
                }

                entryValueBuffer = entryValueBufferMemoryStream.ToArray();
                useB5Header = true;
            }
            else if (localDescriptorList != null)
            {
                LocalDescriptorListElement element = localDescriptorList.Elements.ContainsKey(tableHeader7C.ValueEntriesIndexReference) ? localDescriptorList.Elements[tableHeader7C.ValueEntriesIndexReference] : null;

                if (element == null && tableHeader7C.ValueEntriesIndexReference != 63 && !Util.IsInternalReference(tableHeader7C.ValueEntriesIndexReference))
                {
                    uint elementId = isRecipientsTable ? 1682u : 1649u;

                    element = localDescriptorList.Elements.ContainsKey(elementId) ? localDescriptorList.Elements[elementId] : null;

                    if (element != null && element.SubList != null)
                    {
                        element = element.SubList.Elements.ContainsKey(tableHeader7C.ValueEntriesIndexReference) ? element.SubList.Elements[tableHeader7C.ValueEntriesIndexReference] : null;
                    }
                }
                
                if (element != null)
                {
                    DataStructure dataNode = reader.PstFile.DataIndexTree.GetDataStructure(element.DataStructureId);

                    if (dataNode != null)
                    {
                        entryValueBuffer = Util.GetBuffer(reader, dataNode);
                    }
                }
                else
                {
                    if (tableHeader7C.ValueEntriesIndexReference == 63)
                    {
                        TableIndexItem entryValueIndexItem = index.GetItem(tableHeaderB5.ValueEntriesIndexReference);

                        if (entryValueIndexItem != null && tableBuffer.Length >= entryValueIndexItem.EndOffset)
                        {
                            entryValueBuffer = new byte[entryValueIndexItem.EndOffset - entryValueIndexItem.StartOffset];
                            System.Array.Copy(tableBuffer, entryValueIndexItem.StartOffset, entryValueBuffer, 0, entryValueBuffer.Length);
                        }

                        useB5Header = true;
                    }
                    else
                    {
                        TableIndexItem entryValueIndexItem = index.GetItem(tableHeader7C.ValueEntriesIndexReference);

                        if (entryValueIndexItem != null && tableBuffer.Length >= entryValueIndexItem.EndOffset)
                        {
                            entryValueBuffer = new byte[entryValueIndexItem.EndOffset - entryValueIndexItem.StartOffset];
                            System.Array.Copy(tableBuffer, entryValueIndexItem.StartOffset, entryValueBuffer, 0, entryValueBuffer.Length);
                        }
                    }
                }                
            }
            else
            {
                if (tableHeader7C.ValueEntriesIndexReference == 63)
                {
                        TableIndexItem entryValueIndexItem = index.GetItem(tableHeaderB5.ValueEntriesIndexReference);

                        if (entryValueIndexItem != null && tableBuffer.Length >= entryValueIndexItem.EndOffset)
                        {
                            entryValueBuffer = new byte[entryValueIndexItem.EndOffset - entryValueIndexItem.StartOffset];
                            System.Array.Copy(tableBuffer, entryValueIndexItem.StartOffset, entryValueBuffer, 0, entryValueBuffer.Length);
                        }

                        useB5Header = true;
                }
                else
                {
                    TableIndexItem entryValueIndexItem = index.GetItem(tableHeader7C.ValueEntriesIndexReference);

                    if (entryValueIndexItem != null && tableBuffer.Length >= entryValueIndexItem.EndOffset)
                    {
                        entryValueBuffer = new byte[entryValueIndexItem.EndOffset - entryValueIndexItem.StartOffset];
                        System.Array.Copy(tableBuffer, entryValueIndexItem.StartOffset, entryValueBuffer, 0, entryValueBuffer.Length);
                    }
                }
            }

            if (entryValueBuffer == null)
            {
                return;
            }

            int entriesArraySize = entryValueBuffer.Length / tableHeader7C.ValueArraySize;

            if (useB5Header)
            {
                entriesArraySize = entryValueBuffer.Length / (tableHeaderB5.EntryIdSize + tableHeaderB5.EntryValueSize);
            }

            EntriesArray = new TableEntryList[entriesArraySize];

            for (int m = 0; m < entriesArraySize; m++)
            {
                EntriesArray[m] = new TableEntryList();

                for (int i = 0; i < tableHeader7C.EntryDefinitions.Count; i++)
                {
                    Table7CEntryDefinition entryDefinition = tableHeader7C.EntryDefinitions[i];

                    TableEntry entry = new TableEntry(reader.PstFile.Encoding);
                    entry.PropertyTag = new PropertyTag(entryDefinition.EntryId, entryDefinition.EntryValueType);

                    int offset = m * tableHeader7C.ValueArraySize + entryDefinition.ValueArrayEntryOffset;

                    if (useB5Header)
                    {
                        offset = m * (tableHeaderB5.EntryIdSize + tableHeaderB5.EntryValueSize);
                    }

                    if (entryValueBuffer.Length < offset)
                    {
                      if (isRecipientsTable)
                        Console.Error.WriteLine("\rWarning: Bad entry Recipients Property Table..skipping entry");
                      else
                        Console.Error.WriteLine("\rWarning: Bad entry in Table 7c..skipping entry");
                    }
                    else if (entry.PropertyTag.Type == PropertyType.Short || entry.PropertyTag.Type == PropertyType.Boolean)
                    {
                        entry.ValueBuffer = new byte[2];
                        System.Array.Copy(entryValueBuffer, offset, entry.ValueBuffer, 0, 2);
                    }
                    else if (entry.PropertyTag.Type == PropertyType.Integer || entry.PropertyTag.Type == PropertyType.Float)
                    {
                        entry.ValueBuffer = new byte[4];
                        System.Array.Copy(entryValueBuffer, offset, entry.ValueBuffer, 0, 4);
                    }
                    else if (!useB5Header && (entry.PropertyTag.Type == PropertyType.Long || entry.PropertyTag.Type == PropertyType.Double || entry.PropertyTag.Type == PropertyType.Currency))
                    {
                        entry.ValueBuffer = new byte[8];
                        System.Array.Copy(entryValueBuffer, offset, entry.ValueBuffer, 0, 8);
                    }
                    else if (entry.PropertyTag.Type == PropertyType.Long || entry.PropertyTag.Type == PropertyType.Double || entry.PropertyTag.Type == PropertyType.Currency ||
                             entry.PropertyTag.Type == PropertyType.ShortArray || entry.PropertyTag.Type == PropertyType.IntegerArray || entry.PropertyTag.Type == PropertyType.FloatArray ||
                             entry.PropertyTag.Type == PropertyType.LongArray || entry.PropertyTag.Type == PropertyType.DoubleArray || entry.PropertyTag.Type == PropertyType.CurrencyArray ||
                             entry.PropertyTag.Type == PropertyType.String || entry.PropertyTag.Type == PropertyType.String8 || entry.PropertyTag.Type == PropertyType.StringArray || entry.PropertyTag.Type == PropertyType.String8Array ||
                             entry.PropertyTag.Type == PropertyType.Binary || entry.PropertyTag.Type == PropertyType.BinaryArray || entry.PropertyTag.Type == PropertyType.Object ||
                             entry.PropertyTag.Type == PropertyType.Guid || entry.PropertyTag.Type == PropertyType.GuidArray || entry.PropertyTag.Type == PropertyType.ApplicationTime || entry.PropertyTag.Type == PropertyType.SystemTime)
                    {
                        uint valueReference = 0;

                        if (useB5Header && tableHeaderB5.EntryValueSize == 2)
                        {
                            valueReference = BitConverter.ToUInt16(entryValueBuffer, offset);
                        }
                        else
                        {
                            valueReference = BitConverter.ToUInt32(entryValueBuffer, offset);
                        }

                        if (valueReference == 0)
                        {
                            continue;
                        }

                        if (Util.IsInternalReference(valueReference))
                        {
                            TableIndexItem indexItem = index.GetItem(valueReference);

                            if (indexItem != null)
                            {
                                if (indexItem.EndOffset > indexItem.StartOffset)
                                {
                                    entry.ValueBuffer = new byte[indexItem.EndOffset - indexItem.StartOffset];

                                    if (tableBuffer.Length >= indexItem.StartOffset + entry.ValueBuffer.Length)
                                    {
                                        System.Array.Copy(tableBuffer, indexItem.StartOffset, entry.ValueBuffer, 0, entry.ValueBuffer.Length);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (localDescriptorList != null)
                            {
                                LocalDescriptorListElement element = localDescriptorList.Elements.ContainsKey(valueReference) ? localDescriptorList.Elements[valueReference] : null;

                                if (element == null)
                                {
                                    LocalDescriptorListElement recipientElement = localDescriptorList.Elements.ContainsKey((uint)1682) ? localDescriptorList.Elements[(uint)1682] : null;

                                    if (recipientElement != null && recipientElement.SubList != null)
                                    {
                                        element = recipientElement.SubList.Elements.ContainsKey(valueReference) ? recipientElement.SubList.Elements[valueReference] : null;
                                    }
                                }

                                if (element != null)
                                {
                                    DataStructure dataNode = reader.PstFile.DataIndexTree.GetDataStructure(element.DataStructureId);

                                    if (dataNode != null)
                                    {
                                        entry.ValueBuffer = Util.GetBuffer(reader, dataNode);
                                    }
                                }
                            }
                        }
                    }

                    EntriesArray[m].Add(entry);
                }
            }

            //Try to recover recipients table if email address is null in all entries
            int emptyEmailAddressCount = 0;

            for (int k = 0; k < EntriesArray.Length; k++)
            {
                if(EntriesArray[k][MapiPropertyTag.PR_EMAIL_ADDRESS] != null && EntriesArray[k][MapiPropertyTag.PR_EMAIL_ADDRESS].ValueBuffer == null)
                {
                    emptyEmailAddressCount++;
                }
                else if (EntriesArray[k][MapiPropertyTag.PR_EMAIL_ADDRESS] == null && EntriesArray[k][MapiPropertyTag.PR_RECIPIENT_TYPE] != null)
                {
                    emptyEmailAddressCount++;
                }
            }

            if (emptyEmailAddressCount > (int)(EntriesArray.Length / 2))
            {
                TableInfo tableInfo = new TableInfo(tableHeader7CBuffer);
                IDictionary<uint, byte[]> rowTable = new Dictionary<uint, byte[]>();
                int rowLength = tableInfo.EndOffsetCellExistenceBlock;

                IList<TableRowId> rowIds = new List<TableRowId>();

                for (int i = 0; i < entryValueBuffer.Length; i += entryValueBuffer.Length / entriesArraySize)
                {
                    byte[] rowIdBuffer = new byte[entryValueBuffer.Length / entriesArraySize];

                    if (entryValueBuffer.Length >= i + rowIdBuffer.Length)
                    {
                        System.Array.Copy(entryValueBuffer, i, rowIdBuffer, 0, rowIdBuffer.Length);

                        TableRowId rowId = new TableRowId(rowIdBuffer);
                        rowIds.Add(rowId);
                    }
                }

                if (localDescriptorList != null)
                {
                    LocalDescriptorListElement element = localDescriptorList.Elements.ContainsKey((uint)0x692) ? localDescriptorList.Elements[(uint)0x692] : null;

                    if (element != null)
                    {
                        if (element.SubList == null || !element.SubList.Elements.ContainsKey((uint)63) || element.SubList.Elements[(uint)63] == null)
                        {
                            return;
                        }

                        LocalDescriptorListElement subElement = element.SubList.Elements[(uint)63];
                        DataStructure dataNode = reader.PstFile.DataIndexTree.GetDataStructure(subElement.DataStructureId);

                        if (dataNode != null)
                        {
                            byte[] rowMatrix = Util.GetBuffer(reader, dataNode);

                            int tableSize = reader.PstFile.Is64Bit ? 8176 : 8180;
                            int tableCount = 1;

                            for (int r = 0; r <= rowMatrix.Length - rowLength; r += rowLength)
                            {
                                if (((tableSize * tableCount) - r) < rowLength)
                                {
                                    r = tableSize * tableCount++;
                                }

                                byte[] rowBuffer = new byte[rowLength];

                                System.Array.Copy(rowMatrix, r, rowBuffer, 0, rowBuffer.Length);

                                uint dwRowId = BitConverter.ToUInt32(rowBuffer, 0);

                                if(!rowTable.ContainsKey(dwRowId))
                                {
                                    rowTable.Add(dwRowId, rowBuffer);
                                }
                                else
                                {
                                    //Console.WriteLine("Warning: existing row key:" + dwRowId);
                                }
                            }
                            
                            int cellExistenceBlockLength = tableInfo.EndOffsetCellExistenceBlock - tableInfo.EndOffset1ByteDataValue;
                            EntriesArray = new TableEntryList[rowIds.Count];

                            for (int i = 0; i < rowIds.Count; i++)
                            {
                                EntriesArray[i] = new TableEntryList();

                                byte[] rowBuffer = rowTable.ContainsKey(rowIds[i].Id) ? rowTable[rowIds[i].Id] : null;

                                if (rowBuffer != null)
                                {
                                    for (int c = 0; c < tableInfo.Columns.Count; c++)
                                    {
                                        TableColumnDescription columnDescription = tableInfo.Columns[c];

                                        byte[] cellExistenceBlock = new byte[cellExistenceBlockLength];

                                        System.Array.Copy(rowBuffer, rowBuffer.Length - cellExistenceBlockLength, cellExistenceBlock, 0, cellExistenceBlock.Length);

                                        int cellExistenceBlockValue = cellExistenceBlock[columnDescription.IBit / 8] & (1 << (7 - (columnDescription.IBit % 8)));

                                        if (cellExistenceBlockValue != 0) //true
                                        {
                                            PropertyTag propertyTag = new PropertyTag(columnDescription.Tag);
                                            TableEntry entry = new TableEntry(reader.PstFile.Encoding);

                                            byte[] valueBuffer = new byte[columnDescription.CountBytes];
                                            System.Array.Copy(rowBuffer, columnDescription.Offset, valueBuffer, 0, valueBuffer.Length);

                                            if (propertyTag.Type == PropertyType.Boolean)
                                            {
                                                if (valueBuffer[0] != 0) //true
                                                {
                                                    valueBuffer = BitConverter.GetBytes((short)1);
                                                }
                                                else
                                                {
                                                    valueBuffer = BitConverter.GetBytes((short)0);
                                                }
                                            }
                                            else if (propertyTag.Type == PropertyType.Long || propertyTag.Type == PropertyType.Double || propertyTag.Type == PropertyType.Currency ||
                                                    propertyTag.Type == PropertyType.ApplicationTime || propertyTag.Type == PropertyType.SystemTime || propertyTag.Type == PropertyType.Guid)
                                            {
                                                HeapId heapId = new HeapId(valueBuffer);

                                                TableIndexItem indexItem = index.Items.ContainsKey((int)heapId.Index) ? index.Items[(int)heapId.Index] : null;

                                                if (indexItem != null)
                                                {
                                                    valueBuffer = new byte[indexItem.EndOffset - indexItem.StartOffset];
                                                    System.Array.Copy(tableBuffer, indexItem.StartOffset, valueBuffer, 0, valueBuffer.Length);
                                                }
                                            }
                                            else if (propertyTag.Type == PropertyType.String || propertyTag.Type == PropertyType.String8 || propertyTag.Type == PropertyType.Binary || propertyTag.Type == PropertyType.Object ||
                                                    propertyTag.Type == PropertyType.ShortArray || propertyTag.Type == PropertyType.IntegerArray || propertyTag.Type == PropertyType.FloatArray ||
                                                    propertyTag.Type == PropertyType.LongArray || propertyTag.Type == PropertyType.DoubleArray || propertyTag.Type == PropertyType.CurrencyArray ||
                                                    propertyTag.Type == PropertyType.StringArray || propertyTag.Type == PropertyType.String8Array || propertyTag.Type == PropertyType.BinaryArray || propertyTag.Type == PropertyType.GuidArray)
                                            {
                                                HeapId heapId = new HeapId(valueBuffer);

                                                if (heapId.Index > 0 && heapId.BlockIndex == 0)
                                                {
                                                    TableIndexItem indexItem = index.Items.ContainsKey((int)heapId.Index) ? index.Items[(int)heapId.Index] : null;

                                                    if (indexItem != null)
                                                    {
                                                        valueBuffer = new byte[indexItem.EndOffset - indexItem.StartOffset];
                                                        System.Array.Copy(tableBuffer, indexItem.StartOffset, valueBuffer, 0, valueBuffer.Length);
                                                    }
                                                }
                                                else if (heapId.Index > 0 && heapId.BlockIndex > 0)
                                                {
                                                    if (index.SecondItems.Length >= heapId.BlockIndex)
                                                    {
                                                        IDictionary<int, TableIndexItem> secondIndexTable = index.SecondItems[heapId.BlockIndex - 1];

                                                        if (secondIndexTable != null)
                                                        {
                                                            TableIndexItem indexItem = secondIndexTable.ContainsKey((int)heapId.Index) ? secondIndexTable[(int)heapId.Index] : null;

                                                            if (indexItem != null)
                                                            {
                                                                valueBuffer = new byte[indexItem.EndOffset - indexItem.StartOffset];
                                                                System.Array.Copy(tableBuffer, indexItem.StartOffset, valueBuffer, 0, valueBuffer.Length);
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            entry.PropertyTag = propertyTag;
                                            entry.ValueBuffer = valueBuffer;

                                            EntriesArray[i].Add(entry);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
