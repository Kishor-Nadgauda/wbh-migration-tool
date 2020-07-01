using System;

namespace Independentsoft.Pst
{
    internal class LocalDescriptorListElement
    {  
        private uint id;
        private ulong dataStructureId;
        private ulong subListId;
        private LocalDescriptorList subList;

        internal LocalDescriptorListElement()
        {
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

        internal ulong DataStructureId
        {
            get
            {
                return dataStructureId;
            }
            set
            {
                dataStructureId = value;
            }
        }

        internal ulong SubListId
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

        internal LocalDescriptorList SubList
        {
            get
            {
                return subList;
            }
            set
            {
                subList = value;
            }
        }
       
        #endregion
    }
}
