using System;

namespace Independentsoft.Pst
{
    internal class DataStructure : IndexNodeItem
    {
        private ulong id;
        private ulong offset;
        private ushort size;
        private ushort flag;

        internal DataStructure()
        {
        }

        internal DataStructure(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {

        }

        internal DataStructure Clone()
        {
            DataStructure clone = new DataStructure();

            clone.id = this.id;
            clone.offset = this.offset;
            clone.size = this.size;
            clone.flag = this.flag;

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

        internal ushort Size
        {
            get
            {
                return size;
            }
            set
            {
                size = value;
            }
        }

        internal ushort Flag
        {
            get
            {
                return flag;
            }
            set
            {
                flag = value;
            }
        }
        
        #endregion
    }
}
