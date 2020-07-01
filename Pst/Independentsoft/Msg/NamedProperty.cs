using System;

namespace Independentsoft.Msg
{
    internal class NamedProperty
    {
        private uint id;
        private string name;
        private byte[] guid;
        private NamedPropertyType type = NamedPropertyType.String;

        public NamedProperty()
        {
        }

        #region Properties

        public uint Id
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

        public string Name
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

        public byte[] Guid
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

        public NamedPropertyType Type
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
