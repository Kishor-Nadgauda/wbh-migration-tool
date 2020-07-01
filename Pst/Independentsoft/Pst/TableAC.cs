using System;
using System.IO;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    internal class TableAC : Table
    {
        internal TableAC()
        {
        }

        internal TableAC(PstFileReader reader, byte[] tableBuffer, LocalDescriptorList localDescriptorList, DataStructure tableDataNode)
        {
            Parse(reader, tableBuffer, localDescriptorList, tableDataNode);
        }

        private void Parse(PstFileReader reader, byte[] tableBuffer, LocalDescriptorList localDescriptorList, DataStructure tableDataNode)
        {
            ushort indexOffset = BitConverter.ToUInt16(tableBuffer, 0);
            ushort tableType = BitConverter.ToUInt16(tableBuffer, 2);
            uint tableValueReference = BitConverter.ToUInt32(tableBuffer, 4);

            TableIndex index = new TableIndex(tableBuffer, indexOffset, reader.PstFile.Is64Bit);

            TableIndexItem tableHeaderACIndexItem = index.GetItem(tableValueReference);

            byte[] tableHeaderACBuffer = new byte[tableHeaderACIndexItem.EndOffset - tableHeaderACIndexItem.StartOffset];
            System.Array.Copy(tableBuffer, tableHeaderACIndexItem.StartOffset, tableHeaderACBuffer, 0, tableHeaderACBuffer.Length);

            TableHeaderAC tableHeaderAC = new TableHeaderAC(tableHeaderACBuffer);

            if (tableHeaderAC.Type != 0xAC)
            {
                return;
            }

            byte[] entryDefinitionsBuffer = null;

            if (tableHeaderAC.TableEntryDefinitionReference < 0x8000)
            {
                  TableIndexItem tableEntryDefinitionIndexItem = index.GetItem(tableHeaderAC.TableEntryDefinitionReference);

                  entryDefinitionsBuffer = new byte[tableEntryDefinitionIndexItem.EndOffset - tableEntryDefinitionIndexItem.StartOffset];
                  System.Array.Copy(tableBuffer, tableEntryDefinitionIndexItem.StartOffset, entryDefinitionsBuffer, 0, entryDefinitionsBuffer.Length);
            }
            else
            {
                if (localDescriptorList != null)
                {
                    LocalDescriptorListElement element = localDescriptorList.Elements.ContainsKey(tableHeaderAC.TableEntryDefinitionReference) ? localDescriptorList.Elements[tableHeaderAC.TableEntryDefinitionReference] : null;

                    if (element != null)
                    {
                        DataStructure dataNode = reader.PstFile.DataIndexTree.GetDataStructure(element.DataStructureId);

                        if (dataNode != null)
                        {
                           entryDefinitionsBuffer = Util.GetBuffer(reader, dataNode);
                        }
                    }
                }
            }

            if (entryDefinitionsBuffer == null)
            {
                return;
            }

            IList<TableACEntryDefinition> entryDefinitions = new List<TableACEntryDefinition>();

            for (int i = 0; i < tableHeaderAC.EntryDefinitionCount; i++)
            {
                TableACEntryDefinition entry = new TableACEntryDefinition();

                entry.EntryValueType = BitConverter.ToUInt16(entryDefinitionsBuffer, i * 16);
                entry.EntryId = BitConverter.ToUInt16(entryDefinitionsBuffer, i * 16 + 2);
                entry.ValueArrayEntryOffset = BitConverter.ToUInt16(entryDefinitionsBuffer, i * 16 + 4);
                entry.ValueArrayEntrySize = BitConverter.ToUInt16(entryDefinitionsBuffer, i * 16 + 6);
                entry.ValueArrayEntryNumber = BitConverter.ToUInt16(entryDefinitionsBuffer, i * 16 + 8);
                entry.DescriptorId = BitConverter.ToUInt32(entryDefinitionsBuffer, i * 16 + 12);

                entryDefinitions.Add(entry);
            }

            TableIndexItem tableHeaderB5IndexItem = index.GetItem(tableHeaderAC.B5HeaderReference);

            byte[] tableHeaderB5Buffer = new byte[8];
            System.Array.Copy(tableBuffer, tableHeaderB5IndexItem.StartOffset, tableHeaderB5Buffer, 0, 8);

            TableHeaderB5 tableHeaderB5 = new TableHeaderB5(tableHeaderB5Buffer);

            if (tableHeaderB5.Type != 0xB5)
            {
                return;
            }

            byte[] entryValueBuffer = null;

            if (tableHeaderAC.ValueEntriesIndexReference == 0)
            {
                return;
            }

            if (localDescriptorList != null)
            {
                LocalDescriptorListElement element = localDescriptorList.Elements.ContainsKey(tableHeaderAC.ValueEntriesIndexReference) ? localDescriptorList.Elements[tableHeaderAC.ValueEntriesIndexReference] : null;

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
                    TableIndexItem entryValueIndexItem = index.GetItem(tableHeaderAC.ValueEntriesIndexReference);
                    
                    entryValueBuffer = new byte[entryValueIndexItem.EndOffset - entryValueIndexItem.StartOffset];
                    System.Array.Copy(tableBuffer, entryValueIndexItem.StartOffset, entryValueBuffer, 0, entryValueBuffer.Length);
                }
            }

            if (entryValueBuffer == null)
            {
                return;
            }

            for (int i = 0; i < entryDefinitions.Count; i++)
            {
                TableACEntryDefinition entryDefinition = entryDefinitions[i];

                TableEntry entry = new TableEntry(reader.PstFile.Encoding);
                entry.PropertyTag = new PropertyTag(entryDefinition.EntryId, entryDefinition.EntryValueType);     
     
                if (entry.PropertyTag.Type == PropertyType.Short || entry.PropertyTag.Type == PropertyType.Boolean)
                {
                    entry.ValueBuffer = new byte[2];
                    System.Array.Copy(entryValueBuffer, entryDefinition.ValueArrayEntryOffset, entry.ValueBuffer, 0, 2);
                }
                else if (entry.PropertyTag.Type == PropertyType.Integer || entry.PropertyTag.Type == PropertyType.Float)
                {
                    entry.ValueBuffer = new byte[4];
                    System.Array.Copy(entryValueBuffer, entryDefinition.ValueArrayEntryOffset, entry.ValueBuffer, 0, 4);
                }
                else if (entry.PropertyTag.Type == PropertyType.Long || entry.PropertyTag.Type == PropertyType.Double || entry.PropertyTag.Type == PropertyType.Currency || 
                         entry.PropertyTag.Type == PropertyType.ApplicationTime || entry.PropertyTag.Type == PropertyType.SystemTime)
                {
                    entry.ValueBuffer = new byte[8];
                    System.Array.Copy(entryValueBuffer, entryDefinition.ValueArrayEntryOffset, entry.ValueBuffer, 0, 8);
                }
                else if (entry.PropertyTag.Type == PropertyType.ShortArray || entry.PropertyTag.Type == PropertyType.IntegerArray || entry.PropertyTag.Type == PropertyType.FloatArray ||
                         entry.PropertyTag.Type == PropertyType.LongArray || entry.PropertyTag.Type == PropertyType.DoubleArray || entry.PropertyTag.Type == PropertyType.CurrencyArray ||
                         entry.PropertyTag.Type == PropertyType.String || entry.PropertyTag.Type == PropertyType.String8 || entry.PropertyTag.Type == PropertyType.StringArray || entry.PropertyTag.Type == PropertyType.String8Array ||
                         entry.PropertyTag.Type == PropertyType.Object || entry.PropertyTag.Type == PropertyType.Binary || entry.PropertyTag.Type == PropertyType.BinaryArray || entry.PropertyTag.Type == PropertyType.Guid || entry.PropertyTag.Type == PropertyType.GuidArray)
                { 
                    if (entryDefinition.DescriptorId > 0)
                    {
                        if (localDescriptorList != null)
                        {
                            LocalDescriptorListElement element = localDescriptorList.Elements.ContainsKey(entryDefinition.DescriptorId) ? localDescriptorList.Elements[entryDefinition.DescriptorId] : null;

                            if (element != null)
                            {
                                DataStructure dataNode = reader.PstFile.DataIndexTree.GetDataStructure(element.DataStructureId);

                                if (dataNode != null)
                                {
                                    byte[] valueBuffer = Util.GetBuffer(reader, dataNode);

                                    if (valueBuffer != null && valueBuffer.Length > 3)
                                    {
                                        ushort tableA5Type = BitConverter.ToUInt16(valueBuffer, 2);

                                        if (tableA5Type == 0xA5)
                                        {
                                            TableA5 tableA5 = new TableA5(reader, valueBuffer);
                                            valueBuffer = tableA5.GetValue(entryDefinition.ValueArrayEntryOffset);
                                        }
                                    }
                                    
                                    entry.ValueBuffer = valueBuffer;
                                }
                            }
                        }
                    }
                    else
                    {
                        uint valueReference = BitConverter.ToUInt32(entryValueBuffer, entryDefinition.ValueArrayEntryOffset);

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

                                if (element != null)
                                {
                                    DataStructure dataNode = reader.PstFile.DataIndexTree.GetDataStructure(element.DataStructureId);

                                    if (dataNode != null)
                                    {
                                        byte[] valueBuffer = Util.GetBuffer(reader, dataNode);

                                        if (valueBuffer != null && valueBuffer.Length > 3)
                                        {
                                            ushort tableA5Type = BitConverter.ToUInt16(valueBuffer, 2);

                                            if (tableA5Type == 0xA5)
                                            {
                                                TableA5 tableA5 = new TableA5(reader, valueBuffer);
                                                valueBuffer = tableA5.GetValue(entryDefinition.ValueArrayEntryOffset);
                                            }
                                        }
                                        
                                        entry.ValueBuffer = valueBuffer;
                                    }
                                }
                            }
                        }
                    }
                }

                Entries.Add(entry);
            }
        }
    }
}
