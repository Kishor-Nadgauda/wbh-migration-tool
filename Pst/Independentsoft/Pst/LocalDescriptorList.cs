using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    internal class LocalDescriptorList
    {
        private byte type;
        private byte level;
        private ushort elementCount;
        private IDictionary<uint, LocalDescriptorListElement> elements = new Dictionary<uint, LocalDescriptorListElement>();

        internal LocalDescriptorList()
        {
        }

        internal LocalDescriptorListElement GetFirstElement()
        {
            foreach (LocalDescriptorListElement element in elements.Values)
            {
                if(element.SubList != null)
                {
                    return element;
                }
            }

            return null;
        }

        #region Properties

        internal byte Type
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

        internal ushort ElementCount
        {
            get
            {
                return elementCount;
            }
            set
            {
                elementCount = value;
            }
        }

        internal IDictionary<uint, LocalDescriptorListElement> Elements
        {
            get
            {
                return elements;
            }
        }

        #endregion
    }
}
