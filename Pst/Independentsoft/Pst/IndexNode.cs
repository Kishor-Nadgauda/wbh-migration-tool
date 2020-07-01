using System;
using System.IO;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    internal class IndexNode
    {
        private IndexNode parent;
        private IList<IndexNode> children = new List<IndexNode>();

        private IList<IndexNodeItem> items = new List<IndexNodeItem>();
        private byte itemCount;
        private byte maximumItemCount;
        private byte itemSize;
        private byte level;
        private ushort type;
        private ulong backPointer;
        private uint crc;

        internal IndexNode()
        {
        }

        internal IndexNode(PstFileReader reader, ulong offset, IndexTree tree)
        {
            Parse(reader, offset, tree);
        }

        private void Parse(PstFileReader reader, ulong offset, IndexTree tree)
        {
		  // PRGX: make sure offset is valid
          if (offset > (ulong)reader.BaseStream.Length)
          {
            throw new DataStructureException("Parse - offset beyond length of stream");
          }
            reader.BaseStream.Position = (long)offset;

            if (!reader.PstFile.Is64Bit)
            {
                //32bit
                byte[] itemsBuffer = reader.ReadBytes(496);

                this.itemCount = reader.ReadByte();
                this.maximumItemCount = reader.ReadByte();
                this.itemSize = reader.ReadByte();
                this.level = reader.ReadByte();
                this.type = reader.ReadUInt16();
                this.backPointer = reader.ReadUInt32();
                this.crc = reader.ReadUInt32();

                if (type == 0x8080 && level > 0)
                {
                    for (int i = 0; i < itemCount; i++)
                    {
                        byte[] itemBuffer = new byte[12];
                        System.Array.Copy(itemsBuffer, i * 12, itemBuffer, 0, 12);

                        DataStructureIndexNodeItem item = new DataStructureIndexNodeItem();

                        item.Id = BitConverter.ToUInt32(itemBuffer, 0);
                        item.BackPointer = BitConverter.ToUInt32(itemBuffer, 4);
                        item.Offset = BitConverter.ToUInt32(itemBuffer, 8);

                        items.Add(item);

                        IndexNode childIndexNode = new IndexNode(reader, item.Offset, tree);
                        childIndexNode.Parent = this;

                        this.children.Add(childIndexNode);
                    }
                }
                else if (type == 0x8080 && level == 0)
                {
                    for (int i = 0; i < itemCount; i++)
                    {
                        byte[] itemBuffer = new byte[12];
                        System.Array.Copy(itemsBuffer, i * 12, itemBuffer, 0, 12);

                        DataStructure item = new DataStructure(itemBuffer);

                        item.Id = BitConverter.ToUInt32(itemBuffer, 0);
                        item.Offset = BitConverter.ToUInt32(itemBuffer, 4);
                        item.Size = BitConverter.ToUInt16(itemBuffer, 8);
                        item.Flag = BitConverter.ToUInt16(itemBuffer, 10);

                        items.Add(item);

                        tree.Nodes.Add(item.Id, item);
                    }
                }
                else if (type == 0x8181 && level > 0)
                {
                    for (int i = 0; i < itemCount; i++)
                    {
                        byte[] itemBuffer = new byte[12];
                        System.Array.Copy(itemsBuffer, i * 12, itemBuffer, 0, 12);

                        ItemDescriptorIndexNodeItem item = new ItemDescriptorIndexNodeItem();

                        item.Id = BitConverter.ToUInt32(itemBuffer, 0);
                        item.BackPointer = BitConverter.ToUInt32(itemBuffer, 4);
                        item.Offset = BitConverter.ToUInt32(itemBuffer, 8);

                        items.Add(item);

                        IndexNode childIndexNode = new IndexNode(reader, item.Offset, tree);
                        childIndexNode.Parent = this;

                        this.children.Add(childIndexNode);
                    }
                }
                else if (type == 0x8181 && level == 0)
                {
                    for (int i = 0; i < itemCount; i++)
                    {
                        byte[] itemBuffer = new byte[16];
                        System.Array.Copy(itemsBuffer, i * 16, itemBuffer, 0, 16);

                        ItemDescriptor item = new ItemDescriptor();

                        item.Id = BitConverter.ToUInt32(itemBuffer, 0);
                        item.TableId = BitConverter.ToUInt32(itemBuffer, 4);
                        item.LocalDescriptorListId = BitConverter.ToUInt32(itemBuffer, 8);
                        item.ParentItemDescriptorId = BitConverter.ToUInt32(itemBuffer, 12);

                        items.Add(item);

                        tree.Nodes.Add(item.Id, item);

                        //root folder is parent to self. 
                        if (item.Id == 290)
                        {
                            item.ParentItemDescriptorId = 0;
                        }
                    }
                }
                else
                {
                    throw new ApplicationException("Unknown IndexNode type!");
                }
            }
            else
            {
                //64bit
                byte[] itemsBuffer = reader.ReadBytes(488);

                this.itemCount = reader.ReadByte();
                this.maximumItemCount = reader.ReadByte();
                this.itemSize = reader.ReadByte();
                this.level = reader.ReadByte();
                uint emptyValue = reader.ReadUInt32();
                this.type = reader.ReadUInt16();
                ushort unknow1 = reader.ReadUInt16();
                this.crc = reader.ReadUInt32();
                this.backPointer = reader.ReadUInt64();

                if (type == 0x8080 && level > 0)
                {
                    for (int i = 0; i < itemCount; i++)
                    {
                        byte[] itemBuffer = new byte[24];
                        System.Array.Copy(itemsBuffer, i * 24, itemBuffer, 0, 24);

                        DataStructureIndexNodeItem item = new DataStructureIndexNodeItem();

                        item.Id = BitConverter.ToUInt64(itemBuffer, 0);
                        item.BackPointer = BitConverter.ToUInt64(itemBuffer, 8);
                        item.Offset = BitConverter.ToUInt64(itemBuffer, 16);

                        // PRGX: offset from start of file is only valid when greater than zero
                    	if (item.Offset > 0 && item.Offset < (ulong)reader.BaseStream.Length)
                    	{
							items.Add(item);

                        	IndexNode childIndexNode = new IndexNode(reader, item.Offset, tree);
                        	childIndexNode.Parent = this;

                        	this.children.Add(childIndexNode);
						}
                    }
                }
                else if (type == 0x8080 && level == 0)
                {
                    for (int i = 0; i < itemCount; i++)
                    {
                        byte[] itemBuffer = new byte[24];
                        System.Array.Copy(itemsBuffer, i * 24, itemBuffer, 0, 24);

                        DataStructure item = new DataStructure(itemBuffer);

                        item.Id = BitConverter.ToUInt64(itemBuffer, 0);
                        item.Offset = BitConverter.ToUInt64(itemBuffer, 8);
                        item.Size = BitConverter.ToUInt16(itemBuffer, 16);
                        item.Flag = BitConverter.ToUInt16(itemBuffer, 18);

                    	// PRGX: offset from start of file is only valid when greater than zero
                    	if (item.Offset > 0 && !tree.Nodes.ContainsKey(item.Id))
                    	{
                        	items.Add(item);

                        	tree.Nodes.Add(item.Id, item);
						}
                    }
                }
                else if (type == 0x8181 && level > 0)
                {
                    for (int i = 0; i < itemCount; i++)
                    {
                        byte[] itemBuffer = new byte[24];
                        System.Array.Copy(itemsBuffer, i * 24, itemBuffer, 0, 24);

                        ItemDescriptorIndexNodeItem item = new ItemDescriptorIndexNodeItem();

                        item.Id = BitConverter.ToUInt64(itemBuffer, 0);
                        item.BackPointer = BitConverter.ToUInt64(itemBuffer, 8);
                        item.Offset = BitConverter.ToUInt64(itemBuffer, 16);

                    	// PRGX: offset from start of file is only valid when greater than zero
                    	if (item.Offset > 0)
                    	{
                        	items.Add(item);

                        	IndexNode childIndexNode = new IndexNode(reader, item.Offset, tree);
                        	childIndexNode.Parent = this;

                        	this.children.Add(childIndexNode);
						}
                    }
                }
                else if ((type == 0x8181 || type == 0x0) && level == 0)
                {
                    for (int i = 0; i < itemCount; i++)
                    {
                        byte[] itemBuffer = new byte[32];
                        System.Array.Copy(itemsBuffer, i * 32, itemBuffer, 0, 32);

                        ItemDescriptor item = new ItemDescriptor();

                        item.Id = BitConverter.ToUInt64(itemBuffer, 0);
                        item.TableId = BitConverter.ToUInt64(itemBuffer, 8);
                        item.LocalDescriptorListId = BitConverter.ToUInt64(itemBuffer, 16);
                        item.ParentItemDescriptorId = BitConverter.ToUInt32(itemBuffer, 24);

                        if (item.Id > 0 && !tree.Nodes.ContainsKey(item.Id))
                        {
                        	items.Add(item);

                        	tree.Nodes.Add(item.Id, item);

                        	//root folder is parent to self. 
                        	if (item.Id == 290)
                        	{
                            	item.ParentItemDescriptorId = 0;
                        	}
						}
                    }
                }
                else
                {
                    throw new ApplicationException("Unknown IndexNode type: " + type + " level: " + level);
                }
            }
        }

        #region Properties

        internal IndexNode Parent
        {
            get
            {
                return parent;
            }
            set
            {
                parent = value;
            }
        }

        internal IList<IndexNode> Children
        {
            get
            {
                return children;
            }
        }

        internal IList<IndexNodeItem> Items
        {
            get
            {
                return items;
            }
        }

        internal byte ItemCount
        {
            get
            {
                return itemCount;
            }
            set
            {
                itemCount = value;
            }
        }

        internal byte MaximumItemCount
        {
            get
            {
                return maximumItemCount;
            }
            set
            {
                maximumItemCount = value;
            }
        }

        internal byte ItemSize
        {
            get
            {
                return itemSize;
            }
            set
            {
                itemSize = value;
            }
        }

        internal byte Level
        {
            get
            {
                return level;
            }
            set
            {
                level = value;
            }
        }

        internal ushort Type
        {
            get
            {
                return type;
            }
            set
            {
                type = value;
            }
        }

        internal ulong BackPointer
        {
            get
            {
                return backPointer;
            }
            set
            {
                backPointer = value;
            }
        }

        internal uint Crc
        {
            get
            {
                return crc;
            }
            set
            {
                crc = value;
            }
        }

        #endregion
    }
}
