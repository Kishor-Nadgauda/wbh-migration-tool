using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    internal class ItemDescriptor : IndexNodeItem
    {  
        private ulong id;
        private ulong tableId;
        private ulong localDescriptorListId;
        private ulong parentItemDescriptorId;

        private IList<ItemDescriptor> children = new List<ItemDescriptor>();

        internal ItemDescriptor()
        {
        }

        internal ItemDescriptor Clone()
        {
            ItemDescriptor clone = new ItemDescriptor();

            clone.id = this.id;
            clone.TableId = this.TableId;
            clone.localDescriptorListId = this.localDescriptorListId;
            clone.parentItemDescriptorId = this.parentItemDescriptorId;

            return clone;
        }

        internal override IndexNodeItem CloneIndexNodeItem()
        {
            return Clone();
        }

        #region Properties

        internal ulong Id
        {
            get
            {
                return id;
            }
            set
            {
                id = value;
            }
        }

        internal ulong TableId
        {
            get
            {
                return tableId;
            }
            set
            {
                tableId = value;
            }
        }

        internal ulong LocalDescriptorListId
        {
            get
            {
                return localDescriptorListId;
            }
            set
            {
                localDescriptorListId = value;
            }
        }

        internal ulong ParentItemDescriptorId
        {
            get
            {
                return parentItemDescriptorId;
            }
            set
            {
                parentItemDescriptorId = value;
            }
        }

        internal IList<ItemDescriptor> Children
        {
            get
            {
                return children;
            }
        }

        #endregion
    }
}
