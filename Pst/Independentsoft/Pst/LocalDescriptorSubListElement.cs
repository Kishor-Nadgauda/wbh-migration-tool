using System;

namespace Independentsoft.Pst
{
    internal class LocalDescriptorSubListElement
    {  
        private uint id;
        private uint subListId;

        internal LocalDescriptorSubListElement()
        {
        }

        internal LocalDescriptorSubListElement(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            this.id = BitConverter.ToUInt32(buffer, 0);
            this.subListId = BitConverter.ToUInt32(buffer, 4);
        }

        internal LocalDescriptorSubListElement Clone()
        {
            LocalDescriptorSubListElement clone = new LocalDescriptorSubListElement();

            clone.id = this.id;
            clone.subListId = this.subListId;

            return clone;
        }

        #region Properties

        internal uint Id
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

        internal uint SubListId
        {
            get
            {
                return subListId;
            }
            set
            {
                subListId = value;
            }
        }

        #endregion
    }
}
