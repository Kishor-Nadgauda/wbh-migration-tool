using System;
using System.IO;

namespace Independentsoft.Pst
{
  internal class TableBC : Table
  {
    internal TableBC()
    {
    }

    internal TableBC(PstFileReader reader, byte[] tableBuffer, LocalDescriptorList localDescriptorList, DataStructure tableDataNode)
    {
      try
      {
        Parse(reader, tableBuffer, localDescriptorList, tableDataNode);
      }
      catch (Exception e)
      {
        System.Diagnostics.Trace.WriteLine(String.Format("\nTableBC parsing error - run ScanPST to correct, attempting to continue:\n{0}", e.Message));
      }
    }

    private void Parse(PstFileReader reader, byte[] tableBuffer, LocalDescriptorList localDescriptorList, DataStructure tableDataNode)
    {
      ushort indexOffset = BitConverter.ToUInt16(tableBuffer, 0);
      ushort tableType = BitConverter.ToUInt16(tableBuffer, 2);
      uint tableValueReference = BitConverter.ToUInt32(tableBuffer, 4);

      TableIndex index = new TableIndex(tableBuffer, indexOffset, reader.PstFile.Is64Bit);

      TableIndexItem tableHeaderB5IndexItem = index.GetItem(tableValueReference);

      if (tableHeaderB5IndexItem == null)
      {
        return;
      }

      byte[] tableHeaderB5Buffer = new byte[8];
      System.Array.Copy(tableBuffer, tableHeaderB5IndexItem.StartOffset, tableHeaderB5Buffer, 0, 8);

      TableHeaderB5 tableHeaderB5 = new TableHeaderB5(tableHeaderB5Buffer);
      BTreeOnHeapHeader bTreeOnHeapHeader = new BTreeOnHeapHeader(tableHeaderB5Buffer);

      if (tableHeaderB5.Type != 0xB5)
      {
        return;
      }

      if (tableHeaderB5.ValueEntriesIndexReference == 0)
      {
        return;
      }

      uint valueEntryReference = tableHeaderB5.ValueEntriesIndexReference;
      byte[] entryValueBuffer = null;

      if (bTreeOnHeapHeader.BIdxLevels == 2)
      {
        //Not implemented
        return;
      }
      else if (bTreeOnHeapHeader.BIdxLevels == 1)
      {
        TableIndexItem entryIndexItem = index.Items[(int)bTreeOnHeapHeader.HidRoot.Index];

        byte[] level1EntryValueBuffer = new byte[entryIndexItem.EndOffset - entryIndexItem.StartOffset];
        System.Array.Copy(tableBuffer, entryIndexItem.StartOffset, level1EntryValueBuffer, 0, level1EntryValueBuffer.Length);

        MemoryStream entryValueBufferMemoryStream = new MemoryStream();

        for (int i = 0; i < level1EntryValueBuffer.Length; i += 6)
        {
          byte[] heapIdBuffer = new byte[4];
          System.Array.Copy(level1EntryValueBuffer, i + 2, heapIdBuffer, 0, 4);

          HeapId heapId = new HeapId(heapIdBuffer);

          TableIndexItem tempEntryIndexItem = index.Items[(int)heapId.Index];

          byte[] tempLevel1EntryValueBuffer = new byte[tempEntryIndexItem.EndOffset - tempEntryIndexItem.StartOffset];
          System.Array.Copy(tableBuffer, tempEntryIndexItem.StartOffset, tempLevel1EntryValueBuffer, 0, tempLevel1EntryValueBuffer.Length);

          entryValueBufferMemoryStream.Write(tempLevel1EntryValueBuffer, 0, tempLevel1EntryValueBuffer.Length);
        }

        entryValueBuffer = entryValueBufferMemoryStream.ToArray();
      }
      else
      {
        TableIndexItem entryIndexItem = index.GetItem(valueEntryReference);

        if (entryIndexItem != null)
        {
          entryValueBuffer = new byte[entryIndexItem.EndOffset - entryIndexItem.StartOffset];
          System.Array.Copy(tableBuffer, entryIndexItem.StartOffset, entryValueBuffer, 0, entryValueBuffer.Length);
        }
      }
      if (entryValueBuffer != null)
      {
        for (int i = 0; i < entryValueBuffer.Length - 2; i += 8)
        {
          ushort tag = BitConverter.ToUInt16(entryValueBuffer, i);
          ushort type = BitConverter.ToUInt16(entryValueBuffer, i + 2);

          TableEntry entry = new TableEntry(reader.PstFile.Encoding);
          entry.PropertyTag = new PropertyTag(tag, type);

          if (entry.PropertyTag.Type == PropertyType.Boolean)
          {
            int intValue = BitConverter.ToInt32(entryValueBuffer, i + 4);

            if (intValue != 0) //true
            {
              entry.ValueBuffer = BitConverter.GetBytes((short)1);
            }
            else
            {
              entry.ValueBuffer = BitConverter.GetBytes((short)0);
            }
          }
          else if (entry.PropertyTag.Type == PropertyType.Short)
          {
            entry.ValueBuffer = new byte[2];
            System.Array.Copy(entryValueBuffer, i + 4, entry.ValueBuffer, 0, 2);
          }
          else if (entry.PropertyTag.Type == PropertyType.Integer || entry.PropertyTag.Type == PropertyType.Float)
          {
            entry.ValueBuffer = new byte[4];
            System.Array.Copy(entryValueBuffer, i + 4, entry.ValueBuffer, 0, 4);
          }
          else if (entry.PropertyTag.Type == PropertyType.Object)
          {
            uint valueReference = BitConverter.ToUInt32(entryValueBuffer, i + 4);

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

                  uint embeddedItemId = BitConverter.ToUInt32(entry.ValueBuffer, 0);

                  if (localDescriptorList != null)
                  {
                    LocalDescriptorListElement element = localDescriptorList.Elements.ContainsKey(embeddedItemId) ? localDescriptorList.Elements[embeddedItemId] : null;

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
                    entry.ValueBuffer = valueBuffer;
                  }
                }
              }
            }
          }
          else if (entry.PropertyTag.Type == PropertyType.Long || entry.PropertyTag.Type == PropertyType.Double || entry.PropertyTag.Type == PropertyType.Currency ||
              entry.PropertyTag.Type == PropertyType.ShortArray || entry.PropertyTag.Type == PropertyType.IntegerArray || entry.PropertyTag.Type == PropertyType.FloatArray ||
              entry.PropertyTag.Type == PropertyType.LongArray || entry.PropertyTag.Type == PropertyType.DoubleArray || entry.PropertyTag.Type == PropertyType.CurrencyArray ||
              entry.PropertyTag.Type == PropertyType.String || entry.PropertyTag.Type == PropertyType.String8 || entry.PropertyTag.Type == PropertyType.StringArray || entry.PropertyTag.Type == PropertyType.String8Array ||
              entry.PropertyTag.Type == PropertyType.Binary || entry.PropertyTag.Type == PropertyType.BinaryArray || entry.PropertyTag.Type == PropertyType.Guid ||
              entry.PropertyTag.Type == PropertyType.GuidArray || entry.PropertyTag.Type == PropertyType.ApplicationTime || entry.PropertyTag.Type == PropertyType.SystemTime)
          {
            uint valueReference = BitConverter.ToUInt32(entryValueBuffer, i + 4);

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
                    entry.ValueBuffer = valueBuffer;
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
}
