using System;

namespace Independentsoft.Pst
{
    internal class Array
    {
        private ushort signature;
        private ushort count;
        private uint size;
        private ulong[] items;

        private bool isValid = true;

        internal Array()
        {
        }

        internal Array(byte[] buffer, bool is64Bit)
        {
            this.signature = BitConverter.ToUInt16(buffer, 0);
            this.count = BitConverter.ToUInt16(buffer, 2);
            this.size = BitConverter.ToUInt32(buffer, 4);

            if (is64Bit)
            {
                if(count * 8 + 8 > buffer.Length)
                {
                    isValid = false;
                }
            }
            else
            {
                if(count * 4 + 8 > buffer.Length)
                {
                    isValid = false;
                }
            }

            if(isValid)
            {
                items = new ulong[count];

                if (is64Bit)
                {
                    if(count * 8 + 8 > buffer.Length)
                    {
                        isValid = false;
                    }

                    for (int i = 0; i < count; i++)
                    {
                        items[i] = BitConverter.ToUInt64(buffer, i*8 + 8);
                    }
                }
                else
                {
                    for (int i = 0; i < count; i++)
                    {
                        items[i] = BitConverter.ToUInt32(buffer, i*4 + 8);
                    }
                }
            }
        }

        #region Properties

        internal ulong[] Items
        {
            get
            {
                return items;
            }
        }

        internal bool IsValid
        {
            get
            {
                return isValid;
            }
        }

        #endregion
    }
}
