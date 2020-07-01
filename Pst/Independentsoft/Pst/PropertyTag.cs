using System;

namespace Independentsoft.Pst
{
    public class PropertyTag
    {
        private ushort tag;
        private PropertyType type = PropertyType.Unspecified;

        public PropertyTag()
        {
        }

        public PropertyTag(long propertyTag)
        {
            ushort tagValue = (ushort)(propertyTag >> 16);
            ushort typeValue = (ushort)(propertyTag & 0xFFFF);

            this.tag = tagValue;
            this.type = EnumUtil.ParsePropertyType(typeValue);
        }

        public PropertyTag(int tag, int type)
        {
            this.tag = (ushort)tag;
            this.type = EnumUtil.ParsePropertyType((ushort)type);
        }

        public PropertyTag(int tag, PropertyType type)
        {
            this.tag = (ushort)tag;
            this.type = type;
        }

        #region Properties

        public int Tag
        {
            get
            {
                return tag;
            }
            set
            {
                tag = (ushort)value;
            }
        }

        public PropertyType Type
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
