using System;

namespace Independentsoft.Pst
{
    internal class NamedProperty
    {
        private uint id;
        private string name;
        private byte[] guid;
        private NamedPropertyType type = NamedPropertyType.String;

        internal NamedProperty()
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

        internal string Name
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }

        internal byte[] Guid
        {
            get
            {
                return guid;
            }
            set
            {
                guid = value;
            }
        }

        internal NamedPropertyType Type
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

        #endregion
    }
}
