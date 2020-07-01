using System;
using System.IO;

namespace Independentsoft.Pst
{
    internal class AllocationTable
    {
        private bool[] table = new bool[3968]; //496 * 8
        private ushort type;
        private uint offset;
        private uint crc;

        internal AllocationTable()
        {
        }

        internal AllocationTable(PstFileReader reader, uint offset)
        {
            Parse(reader, offset);
        }

        private void Parse(PstFileReader reader, uint offset)
        {
            reader.BaseStream.Position = offset;

            uint unknown1 = reader.ReadUInt32();
            byte[] byteTable = reader.ReadBytes(496);
            type = reader.ReadUInt16();
            ushort unknown2 = reader.ReadUInt16();
            this.offset = reader.ReadUInt32();
            crc = reader.ReadUInt32();

            for (int i = 0; i < 496; i++)
            {
                table[i] = (byteTable[i] | 127) == 255 ? true : false;
                table[i + 1] = (byteTable[i] | 191) == 255 ? true : false;
                table[i + 2] = (byteTable[i] | 223) == 255 ? true : false;
                table[i + 3] = (byteTable[i] | 239) == 255 ? true : false;
                table[i + 4] = (byteTable[i] | 247) == 255 ? true : false;
                table[i + 5] = (byteTable[i] | 251) == 255 ? true : false;
                table[i + 6] = (byteTable[i] | 253) == 255 ? true : false;
                table[i + 7] = (byteTable[i] | 254) == 255 ? true : false;
            }
        }

        #region Properties

        internal bool[] Table
        {
            get
            {
                return table;
            }
            set
            {
                table = value;
            }
        }

        internal ushort Type
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

        internal uint Offset
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
