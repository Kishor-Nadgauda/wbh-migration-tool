using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class PstFile.
    /// </summary>
    public class PstFile : IDisposable
    {
        private Stream stream;
        private PstFileReader reader;
        private System.Text.Encoding encoding = System.Text.Encoding.UTF8;

        private bool is64Bit;
        private EncryptionType encryptionType = EncryptionType.None;
        private long size;
        private ulong descriptorIndexBackPointer;
        private ulong descriptorIndex;
        private ulong dataIndexBackPointer;
        private ulong dataIndex;
        private IndexTree descriptorIndexTree;
        private IndexTree dataIndexTree;

        private NameToIdMap nameToIdMap;
        private MessageStore messageStore;
        private Folder root;
        private Folder mailboxRoot;

        private IDictionary<object, long> namedIdToTag = new Dictionary<object, long>();
        private static long attachmentSizeLimit = Int64.MaxValue;

        /// <summary>
        /// Initializes a new instance of the <see cref="PstFile"/> class.
        /// </summary>
        public PstFile()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PstFile"/> class.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        public PstFile(string filePath) : this(filePath, System.Text.Encoding.UTF8)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PstFile"/> class.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="encoding">The encoding.</param>
        public PstFile(string filePath, System.Text.Encoding encoding)
        {
            Open(filePath, encoding);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PstFile"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public PstFile(System.IO.Stream stream) : this(stream, System.Text.Encoding.UTF8)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PstFile"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="encoding">The encoding.</param>
        public PstFile(System.IO.Stream stream, System.Text.Encoding encoding) 
        {
            Open(stream, encoding);
        }

        /// <summary>
        /// Opens the specified file path.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        public void Open(string filePath)
        {
            Open(filePath, System.Text.Encoding.UTF8);
        }

        /// <summary>
        /// Opens the specified file path.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="encoding">The encoding.</param>
        public void Open(string filePath, System.Text.Encoding encoding)
        {
            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);

            Open(fileStream, encoding);
        }

        /// <summary>
        /// Opens the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public void Open(Stream stream)
        {
            Open(stream, System.Text.Encoding.UTF8);
        }

        /// <summary>
        /// Opens the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="encoding">The encoding.</param>
        public void Open(Stream stream, System.Text.Encoding encoding)
        {
            Parse(stream, encoding);
        }

        private void Parse(Stream stream, System.Text.Encoding encoding)
        {
            this.encoding = encoding;
            this.stream = stream;
            this.reader = new PstFileReader(this, stream);

            uint signature = reader.ReadUInt32();//offset 0

            if (signature != 0x4e444221)
            {
                Close();

                throw new FileFormatException("Invalid file format!");
            }

            uint crc32 = reader.ReadUInt32(); //offset 4
            ushort contentType = reader.ReadUInt16(); //offset 8
            ushort dataVersion = reader.ReadUInt16(); //offset 10
            ushort contentVersion = reader.ReadUInt16(); //offset 12
            byte creationPlatform = reader.ReadByte(); //offset 14
            byte accessPlatform = reader.ReadByte(); //offset 15
            uint unknown1 = reader.ReadUInt32(); //offset 16
            uint unknown2 = reader.ReadUInt32(); //offset 20

            if (dataVersion == 0x0015 || dataVersion == 0x0017)
            {
                is64Bit = true;
            }

            if (!is64Bit)
            {
                //32-bit file
                uint unknown3 = reader.ReadUInt32(); //offset 24
                uint unknown4 = reader.ReadUInt32(); //offset 28
                uint unknown5 = reader.ReadUInt32(); //offset 32

                byte[] unknownArray1 = reader.ReadBytes(32); //offset 36
                byte[] unknownArray2 = reader.ReadBytes(32);
                byte[] unknownArray3 = reader.ReadBytes(32);
                byte[] unknownArray4 = reader.ReadBytes(32);

                uint unknown6 = reader.ReadUInt32(); //offset 164

                size = reader.ReadInt32(); //offset 168

                uint lastDataAllocationTableOffset = reader.ReadUInt32(); //offset 172
                uint unknown7 = reader.ReadUInt32(); //offset 176
                uint unknown8 = reader.ReadUInt32(); //offset 180

                descriptorIndexBackPointer = reader.ReadUInt32();
                descriptorIndex = reader.ReadUInt32();
                dataIndexBackPointer = reader.ReadUInt32();
                dataIndex = reader.ReadUInt32();

                uint unknown9 = reader.ReadUInt32(); //offset 200
                byte[] dataAllocationTableMap = reader.ReadBytes(128); //offset 204
                byte[] indexNodeAllocationTableMap = reader.ReadBytes(128); //offset 332
                byte senitinal = reader.ReadByte(); //offset 460

                encryptionType = EnumUtil.ParseEncryptionType(reader.ReadByte()); //offset 461

                //AllocationTable dataAllocationTable = new AllocationTable(reader, 17408); //offset 0x4400
                //AllocationTable indexNodeAllocationTable = new AllocationTable(reader, 17920); //offset 0x4600
            }
            else
            {
                //64-bit file
                ulong unknown3 = reader.ReadUInt64(); //offset 24
                ulong unknown4 = reader.ReadUInt64(); //offset 32
                uint unknown5 = reader.ReadUInt32(); //offset 40

                byte[] unknownArray1 = reader.ReadBytes(32); //offset 36
                byte[] unknownArray2 = reader.ReadBytes(32);
                byte[] unknownArray3 = reader.ReadBytes(32);
                byte[] unknownArray4 = reader.ReadBytes(32);

                uint unknown61 = reader.ReadUInt32(); //offset 172
                uint unknown62 = reader.ReadUInt32(); //offset 176
                uint unknown63 = reader.ReadUInt32(); //offset 180

                size = reader.ReadInt64(); //offset 184

                ulong lastDataAllocationTableOffset = reader.ReadUInt64(); //offset 192
                ulong unknown7 = reader.ReadUInt64(); //offset 200
                ulong unknown8 = reader.ReadUInt64(); //offset 208

                descriptorIndexBackPointer = reader.ReadUInt64(); //offset 216
                descriptorIndex = reader.ReadUInt64(); //offset 224
                dataIndexBackPointer = reader.ReadUInt64(); //offset 232
                dataIndex = reader.ReadUInt64(); //offset 240

                uint unknown9 = reader.ReadUInt32(); //offset 248
                uint unknown10 = reader.ReadUInt32(); //offset 252

                byte[] dataAllocationTableMap = reader.ReadBytes(128); //offset 256
                byte[] indexNodeAllocationTableMap = reader.ReadBytes(128); //offset 384

                byte senitinal = reader.ReadByte(); //offset 512

                encryptionType = EnumUtil.ParseEncryptionType(reader.ReadByte()); //offset 513
            }

            descriptorIndexTree = new IndexTree(reader, descriptorIndex);
            descriptorIndexTree.SetParent(); //set parent folder

            dataIndexTree = new IndexTree(reader, dataIndex);

            ulong nameToIdMapId = 97;
            ulong messageStoreId = 33;
            ulong rootId = 290;
            ulong mailboxRootId = 32802; //32802 or 32834

            Table nameToIdMapTable = null;
            Table messageStoreTable = null;
            Table rootTable = null;
            Table mailboxRootTable = null;

            ItemDescriptor nameToIdMapItemDescriptor = descriptorIndexTree.Nodes.ContainsKey(nameToIdMapId) ? (ItemDescriptor)descriptorIndexTree.Nodes[nameToIdMapId] : null;
            ItemDescriptor messageStoreItemDescriptor = descriptorIndexTree.Nodes.ContainsKey(messageStoreId) ? (ItemDescriptor)descriptorIndexTree.Nodes[messageStoreId] : null;
            ItemDescriptor rootItemDescriptor = descriptorIndexTree.Nodes.ContainsKey(rootId) ? (ItemDescriptor)descriptorIndexTree.Nodes[rootId] : null;
            ItemDescriptor mailboxRootItemDescriptor = descriptorIndexTree.Nodes.ContainsKey(mailboxRootId) ? (ItemDescriptor)descriptorIndexTree.Nodes[mailboxRootId] : null;

            if (mailboxRootItemDescriptor == null)
            {
                mailboxRootItemDescriptor = descriptorIndexTree.Nodes.ContainsKey((ulong)32834) ? (ItemDescriptor)descriptorIndexTree.Nodes[(ulong)32834] : null;
            }

            if (nameToIdMapItemDescriptor == null)
            {
              throw new DataStructureException("Unresolvable Data Structure issue with: nameToIdMapItemDescriptor");
            }

            DataStructure nameToIdMapLocalDescriptorListNode = Util.GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, nameToIdMapItemDescriptor.LocalDescriptorListId);
            DataStructure nameToIdMapDataNode = reader.PstFile.DataIndexTree.GetDataStructure(nameToIdMapItemDescriptor.TableId);

            if (nameToIdMapDataNode != null)
            {
                LocalDescriptorList nameToIdMapLocalDescriptorList = Util.GetLocalDescriptorList(reader, nameToIdMapLocalDescriptorListNode);
                nameToIdMapTable = Util.GetTable(reader, nameToIdMapDataNode, nameToIdMapLocalDescriptorList);
            }

            if (messageStoreItemDescriptor == null)
            {
              throw new DataStructureException("Unresolvable Data Structure issue with: messageStoreItemDescriptor");
            }

            DataStructure messageStoreLocalDescriptorListNode = Util.GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, messageStoreItemDescriptor.LocalDescriptorListId);
            DataStructure messageStoreDataNode = reader.PstFile.DataIndexTree.GetDataStructure(messageStoreItemDescriptor.TableId);

            if (messageStoreDataNode != null)
            {
                LocalDescriptorList messageStoreLocalDescriptorList = Util.GetLocalDescriptorList(reader, messageStoreLocalDescriptorListNode);
                messageStoreTable = Util.GetTable(reader, messageStoreDataNode, messageStoreLocalDescriptorList);
            }

            if (rootItemDescriptor == null)
            {
              throw new DataStructureException("Unresolvable Data Structure issue with: rootItemDescriptor");
            }

            DataStructure rootLocalDescriptorListNode = Util.GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, rootItemDescriptor.LocalDescriptorListId);
            DataStructure rootDataNode = reader.PstFile.DataIndexTree.GetDataStructure(rootItemDescriptor.TableId);

            if (rootDataNode != null)
            {
                LocalDescriptorList rootLocalDescriptorList = Util.GetLocalDescriptorList(reader, rootLocalDescriptorListNode);
                rootTable = Util.GetTable(reader, rootDataNode, rootLocalDescriptorList);
            }

            if (mailboxRootItemDescriptor != null)
            {
                DataStructure mailboxRootLocalDescriptorListNode = Util.GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, mailboxRootItemDescriptor.LocalDescriptorListId);
                DataStructure mailboxRootDataNode = reader.PstFile.DataIndexTree.GetDataStructure(mailboxRootItemDescriptor.TableId);

                if (mailboxRootDataNode != null)
                {
                    LocalDescriptorList mailboxRootLocalDescriptorList = Util.GetLocalDescriptorList(reader, mailboxRootLocalDescriptorListNode);
                    mailboxRootTable = Util.GetTable(reader, mailboxRootDataNode, mailboxRootLocalDescriptorList);
                }
            }
            
            if (nameToIdMapTable != null)
            {
                nameToIdMap = new NameToIdMap(nameToIdMapTable);

                FillNamedIdToTag();
            }

            if (messageStoreTable != null)
            {
                messageStore = new MessageStore(messageStoreTable);
            }

            if (rootTable != null)
            {
                root = new Folder(reader, rootItemDescriptor, rootTable);
            }

            if (mailboxRootItemDescriptor != null && mailboxRootTable != null)
            {
                mailboxRoot = new Folder(reader, mailboxRootItemDescriptor, mailboxRootTable);
            }
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Close();
        }

        /// <summary>
        /// Closes this instance.
        /// </summary>
        public void Close()
        {
            if (stream != null)
            {
                stream.Close();
            }
        }

        /// <summary>
        /// Gets the folder.
        /// </summary>
        /// <param name="entryId">The entry identifier.</param>
        /// <returns>Folder.</returns>
        /// <exception cref="System.ArgumentException">
        /// Invalid EntryId.
        /// or
        /// Invalid EntryId.
        /// </exception>
        public Folder GetFolder(byte[] entryId)
        {
            if (entryId != null && entryId.Length == 24)
            {
                byte[] messageStoreRecordKey = new byte[16];
                System.Array.Copy(entryId, 4, messageStoreRecordKey, 0, 16);

                if (!CompareMessageStoreRecordKey(messageStoreRecordKey))
                {
                    throw new ArgumentException("Invalid EntryId.");
                }

                byte[] idBuffer = new byte[4];
                System.Array.Copy(entryId, 20, idBuffer, 0, 4);

                uint folderId = BitConverter.ToUInt32(idBuffer, 0);

                return GetFolder(folderId);
            }
            else
            {
                throw new ArgumentException("Invalid EntryId.");
            }
        }

        /// <summary>
        /// Gets the folder.
        /// </summary>
        /// <param name="id">The identifier.</param>
        /// <returns>Folder.</returns>
        public Folder GetFolder(long id)
        {
            ulong key = (ulong)id;

            ItemDescriptor itemDescriptor = descriptorIndexTree.Nodes.ContainsKey(key) ? (ItemDescriptor)descriptorIndexTree.Nodes[key] : null;

            if (itemDescriptor == null)
            {
                return null;
            }

            Table folderTable = null;
            LocalDescriptorList localDescriptorList = null;

            DataStructure localDescriptorListNode = Util.GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, itemDescriptor.LocalDescriptorListId);
            DataStructure dataNode = reader.PstFile.DataIndexTree.GetDataStructure(itemDescriptor.TableId);

            if (dataNode != null)
            {
                localDescriptorList = Util.GetLocalDescriptorList(reader, localDescriptorListNode);
                folderTable = Util.GetTable(reader, dataNode, localDescriptorList);
            }

            if (folderTable == null)
            {
                return null;
            }

            if (Util.IsFolder(folderTable))
            {
                Folder folder = new Folder(reader, itemDescriptor, folderTable);
                return folder;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Gets the item.
        /// </summary>
        /// <param name="entryId">The entry identifier.</param>
        /// <returns>Item.</returns>
        /// <exception cref="System.ArgumentException">
        /// Invalid EntryId.
        /// or
        /// Invalid EntryId.
        /// </exception>
        public Item GetItem(byte[] entryId)
        {
            if (entryId != null && entryId.Length == 24)
            {            
                byte[] messageStoreRecordKey = new byte[16];
                System.Array.Copy(entryId, 4, messageStoreRecordKey, 0, 16);

                if (!CompareMessageStoreRecordKey(messageStoreRecordKey))
                {
                    throw new ArgumentException("Invalid EntryId.");
                }

                byte[] idBuffer = new byte[4];
                System.Array.Copy(entryId, 20, idBuffer, 0, 4);

                uint itemId = BitConverter.ToUInt32(idBuffer,0);

                return GetItem(itemId);
            }
            else
            {
                throw new ArgumentException("Invalid EntryId.");
            }
        }
        
        public ICollection<ulong> GetItemKeys()
        {
          return descriptorIndexTree.Nodes.Keys;
        }

        /// <summary>
        /// Gets the item.
        /// </summary>
        /// <param name="id">The identifier.</param>
        /// <returns>Item.</returns>
        public Item GetItem(long id)
        {
            ulong key = (ulong)id;

            ItemDescriptor itemDescriptor = descriptorIndexTree.Nodes.ContainsKey(key) ? (ItemDescriptor)descriptorIndexTree.Nodes[key] : null;

            if (itemDescriptor == null)
            {
                return null;
            }

            Table itemTable = null;
            LocalDescriptorList localDescriptorList = null;

            DataStructure localDescriptorListNode = Util.GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, itemDescriptor.LocalDescriptorListId);
            DataStructure dataNode = reader.PstFile.DataIndexTree.GetDataStructure(itemDescriptor.TableId);

            if (dataNode != null)
            {
                localDescriptorList = Util.GetLocalDescriptorList(reader, localDescriptorListNode);
                try
                {
                  itemTable = Util.GetTable(reader, dataNode, localDescriptorList);
                }
                catch (Exception e)
                {
                  System.Diagnostics.Trace.TraceWarning("PstFile.GetItem({0}) Util.GetTable {1}", id, e.Message);
                }
            }

            if (itemTable == null)
            {
                return null;
            }

            if (!Util.IsFolder(itemTable))
            {
                Item item = Util.GetItem(reader, localDescriptorList, itemTable, itemDescriptor.Id, itemDescriptor.ParentItemDescriptorId);
                return item;
            }
            else
            {
                return null;
            }
        }
        public System.Collections.Generic.IEnumerable<Item> GetAllItems()
        {
          foreach (ItemDescriptor itemDescriptor in descriptorIndexTree.Nodes.Values)
          {
            if (itemDescriptor != null)
            {
              Table itemTable = null;
              LocalDescriptorList localDescriptorList = null;

              DataStructure localDescriptorListNode = Util.GetLocalDescriptionListNode(reader.PstFile.DataIndexTree, itemDescriptor.LocalDescriptorListId);
              DataStructure dataNode = reader.PstFile.DataIndexTree.GetDataStructure(itemDescriptor.TableId);

              if (dataNode != null)
              {
                localDescriptorList = Util.GetLocalDescriptorList(reader, localDescriptorListNode);
                itemTable = Util.GetTable(reader, dataNode, localDescriptorList);
              }

              if (itemTable != null)
              {

                if (!Util.IsFolder(itemTable))
                {
                  Item item = Util.GetItem(reader, localDescriptorList, itemTable, itemDescriptor.Id, itemDescriptor.ParentItemDescriptorId);
                  yield return item;
                }
              }
            }
          }
        }

        private bool CompareMessageStoreRecordKey(byte[] recordKey)
        {
            if (recordKey == null && messageStore.RecordKey == null)
            {
                return true;
            }
            else if (recordKey == null || messageStore.RecordKey == null)
            {
                return false;
            }
            else if (recordKey.Length != 16 || messageStore.RecordKey.Length != 16)
            {
                return false;
            }
            else
            {
                for (int i = 0; i < 16; i++)
                {
                    if (recordKey[i] != messageStore.RecordKey[i])
                    {
                        return false;
                    }
                }

                return true;
            }
        }

        private void FillNamedIdToTag()
        {
            foreach (NamedProperty namedProperty in nameToIdMap.Map.Values)
            {
                if (namedProperty.Type == NamedPropertyType.Numerical)
                {
                    try
                    {
                        long tag = nameToIdMap.GetId(namedProperty.Id);

                        namedIdToTag[namedProperty.Id] = tag;
                    }
                    catch
                    {
                    }
                }
                else
                {
                    try
                    {
                        long tag = nameToIdMap.GetId(namedProperty.Name, namedProperty.Guid);

                        string namedPropertyHex = Util.ConvertNamedPropertyToHex(namedProperty.Name, namedProperty.Guid);

                        namedIdToTag[namedPropertyHex] = tag;
                    }
                    catch
                    {
                    }
                }
            }
        }

        #region Properties

        internal IndexTree DescriptorIndexTree
        {
            get
            {
                return descriptorIndexTree;
            }
        }

        internal IndexTree DataIndexTree
        {
            get
            {
                return dataIndexTree;
            }
        }

        internal NameToIdMap NameToIdMap
        {
            get
            {
                return nameToIdMap;
            }
        }

        internal IDictionary<object, long> NamedIdToTag
        {
            get
            {
                return namedIdToTag;
            }
        }

        /// <summary>
        /// Gets a value indicating whether [is64 bit].
        /// </summary>
        /// <value><c>true</c> if [is64 bit]; otherwise, <c>false</c>.</value>
        public bool Is64Bit
        {
            get
            {
                return is64Bit;
            }
        }

        /// <summary>
        /// Gets the size.
        /// </summary>
        /// <value>The size.</value>
        public long Size
        {
            get
            {
                return size;
            }
        }

        /// <summary>
        /// Gets the type of the encryption.
        /// </summary>
        /// <value>The type of the encryption.</value>
        public EncryptionType EncryptionType
        {
            get
            {
                return encryptionType;
            }
        }

        /// <summary>
        /// Gets the message store.
        /// </summary>
        /// <value>The message store.</value>
        public MessageStore MessageStore
        {
            get
            {
                return messageStore;
            }
        }

        /// <summary>
        /// Gets the root.
        /// </summary>
        /// <value>The root.</value>
        public Folder Root
        {
            get
            {
                return root;
            }
        }

        /// <summary>
        /// Gets the mailbox root.
        /// </summary>
        /// <value>The mailbox root.</value>
        public Folder MailboxRoot
        {
            get
            {
                return mailboxRoot;
            }
        }

        /// <summary>
        /// Gets or sets the encoding.
        /// </summary>
        /// <value>The encoding.</value>
        public System.Text.Encoding Encoding
        {
            get
            {
                return encoding;
            }
            set
            {
                encoding = value;
            }
        }

        /// <summary>
        /// Gets or sets the attachment size limit.
        /// </summary>
        /// <value>The attachment size limit.</value>
        public static long AttachmentSizeLimit
        {
            get
            {
                return attachmentSizeLimit;
            }
            set
            {
                attachmentSizeLimit = value;
            }
        }

        #endregion 
    }

  [Serializable]
  public class DataStructureException : Exception
  {
    private Type type;
    private string v;

    public DataStructureException()
    {
    }

    public DataStructureException(string message) : base(message)
    {
    }

    public DataStructureException(string message, Exception innerException) : base(message, innerException)
    {
    }

    protected DataStructureException(SerializationInfo info, StreamingContext context) : base(info, context)
    {
    }
  }
}
