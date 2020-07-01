using System;

namespace Independentsoft.Pst
{
    internal class ItemDescriptorIndexNodeItem : IndexNodeItem
    {  
        private ulong id;
        private ulong backPointer;
        private ulong offset;

        internal ItemDescriptorIndexNodeItem()
        {
        }

        internal ItemDescriptorIndexNodeItem Clone()
        {
            ItemDescriptorIndexNodeItem clone = new ItemDescriptorIndexNodeItem();

            clone.id = this.id;
            clone.backPointer = this.backPointer;
            clone.offset = this.offset;

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

        internal ulong Offset
        {
            get
            {
                return offset;
            }
            set
            {
                offset = value;
            }
        }
       
        #endregion
    }
}
