using System;
using System.IO;

namespace Independentsoft.Msg
{
    internal class Property
    {
        private uint tag;
        private PropertyType type = PropertyType.Integer16;
        private uint size;
        private byte[] value;
        private bool isMandatory;
        private bool isReadable  = true;
        private bool isWriteable = true;

        public Property()
        {
        }

        public Property(byte[] buffer)
        {
            Parse(buffer);
        }

        private void Parse(byte[] buffer)
        {
            tag = BitConverter.ToUInt32(buffer, 0);

            type = EnumUtil.ParsePropertyType(tag);
            
            uint flags = BitConverter.ToUInt32(buffer, 4);

            if (flags == 1)
            {
                isMandatory = true;
            }
            else if (flags == 2)
            {
                isReadable = true;
            }
            else if (flags == 3)
            {
                isMandatory = true;
                isReadable = true;
            }
            else if (flags == 4)
            {
                isWriteable = true;
            }
            else if (flags == 5)
            {
                isMandatory = true;
                isWriteable = true;
            }
            else if (flags == 6)
            {
                isReadable = true;
                isWriteable = true;
            }
            else if (flags == 7)
            {
                isMandatory = true;
                isReadable = true;
                isWriteable = true;
            }

            //fixed length properties
            if(type == PropertyType.Integer16)
            {
                value = new byte[2];

                System.Array.Copy(buffer, 8, value, 0, value.Length);
            }
            else if (type == PropertyType.Integer32)
            {
                value = new byte[4];

                System.Array.Copy(buffer, 8, value, 0, value.Length);
            }
            else if (type == PropertyType.Integer64)
            {
                value = new byte[8];

                System.Array.Copy(buffer, 8, value, 0, value.Length);
            }
            else if (type == PropertyType.Floating32)
            {
                value = new byte[4];
                
                System.Array.Copy(buffer, 8, value, 0, value.Length);
            }
            else if (type == PropertyType.Floating64)
            {
                value = new byte[8];

                System.Array.Copy(buffer, 8, value, 0, value.Length);
            }
            else if (type == PropertyType.Currency)
            {
                value = new byte[8];
                
                System.Array.Copy(buffer, 8, value, 0, value.Length);
            }
            else if (type == PropertyType.FloatingTime)
            {
                value = new byte[8];

                System.Array.Copy(buffer, 8, value, 0, value.Length);
            }
            else if (type == PropertyType.ErrorCode)
            {
                value = new byte[4];

                System.Array.Copy(buffer, 8, value, 0, value.Length);
            }
            else if (type == PropertyType.Boolean)
            {
                value = new byte[2];

                System.Array.Copy(buffer, 8, value, 0, value.Length);
            }
            else if (type == PropertyType.Time)
            {
                value = new byte[8];

                System.Array.Copy(buffer, 8, value, 0, value.Length);
            }
            //multi-value and variable length properties
            else
            {
                size = BitConverter.ToUInt32(buffer, 8);
            }
        }

        public byte[] ToBytes()
        {
            byte[] buffer = new byte[16];
            MemoryStream memoryStream = new MemoryStream(buffer);

            byte[] tagBuffer = BitConverter.GetBytes(tag);
            memoryStream.Write(tagBuffer, 0, 4);

            uint flags = 0;

            if (isMandatory)
            {
                flags += 1;
            }

            if (isReadable)
            {
                flags += 2;
            }

            if (isWriteable)
            {
                flags += 4;
            }

            byte[] flagsBuffer = BitConverter.GetBytes(flags);
            memoryStream.Write(flagsBuffer, 0, 4);

            if (size > 0)
            {
                byte[] sizeBuffer = BitConverter.GetBytes(size);
                memoryStream.Write(sizeBuffer, 0, 4);
            }
            else
            {
                if (value != null && value.Length > 0)
                {
                    memoryStream.Write(value, 0, value.Length);
                }
            }

            return buffer;
        }

        #region Properties

        internal uint Tag
        {
            get
            {
                return tag;
            }
            set
            {
                tag = value;
            }
        }

        internal PropertyType Type
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


        internal bool IsMandatory
        {
            get
            {
                return isMandatory;
            }
            set
            {
                isMandatory = value;
            }
        }

        internal bool IsReadable
        {
            get
            {
                return isReadable;
            }
            set
            {
                isReadable = value;
            }
        }

        internal bool IsWriteable
        {
            get
            {
                return isWriteable;
            }
            set
            {
                isWriteable = value;
            }
        }

        internal uint Size
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

        internal byte[] Value
        {
            get
            {
                return value;
            }
            set
            {
                this.value = value;
            }
        }

        #endregion
    }
}
