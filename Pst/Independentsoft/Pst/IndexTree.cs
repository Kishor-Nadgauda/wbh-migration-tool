using System;
using System.IO;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    internal class IndexTree
    {
        private IndexNode root;
        private IDictionary<ulong, IndexNodeItem> nodes = new Dictionary<ulong, IndexNodeItem>();

        internal IndexTree()
        {
        }

        internal IndexTree(PstFileReader reader, ulong offset)
        {
            Parse(reader, offset);
        }

        private void Parse(PstFileReader reader, ulong offset)
        {
            reader.BaseStream.Position = (long)offset;

            root = new IndexNode(reader, offset, this);
        }

        internal DataStructure GetDataStructure(ulong nodeId)
        {
            nodeId = (nodeId >> 1) << 1;

            if (nodes.ContainsKey(nodeId) && nodes[nodeId] != null)
            {
                return (DataStructure)nodes[nodeId];
            }
            else
            {
                return null;
            }
        }

        internal void SetParent()
        {
            foreach (ulong key in nodes.Keys)
            {
                ItemDescriptor item = (ItemDescriptor)nodes[key];

                if (item.ParentItemDescriptorId > 0 && nodes.ContainsKey(item.ParentItemDescriptorId) && nodes[item.ParentItemDescriptorId] != null)
                {
                    ItemDescriptor parent = (ItemDescriptor)nodes[item.ParentItemDescriptorId];
                    parent.Children.Add(item);
                }
            }
        }

        #region Properties

        internal IndexNode Root
        {
            get
            {
                return root;
            }
        }

        internal IDictionary<ulong, IndexNodeItem> Nodes
        {
            get
            {
                return nodes;
            }
        }

        #endregion
    }
}
