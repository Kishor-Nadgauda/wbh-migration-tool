using System;
using System.IO;

namespace Independentsoft.IO.StructuredStorage
{
    internal class Header
    {
        private byte[] signature = new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
        private byte[] classId = new byte[16];
        private ushort minorVersion = 0x003E;
        private ushort majorVersion = 0x0003;
        private ushort byteOrder = 0xFFFE;
        private ushort sectorSize = 0x0009;
        private ushort miniSectorSize = 0x0006;
        private ushort reserved;
        private uint reserved1;
        private uint directorySectorCount;
        private uint fatSectorCount;
        private uint firstDirectorySector;
        private uint transactionSignature;
        private uint miniStreamMaxSize = 4096;
        private uint firstMiniFatSector;
        private uint miniFatSectorCount;
        private uint firstDifatSector;
        private uint difatSectorCount;
        private uint[] difat = new uint[109];

        internal Header()
        {
        }

        internal Header(BinaryReader reader)
        {
            Parse(reader);
        }

        private void Parse(BinaryReader reader)
        {
            byte[] testSignature = reader.ReadBytes(8);

            for(int i=0; i < 8; i++)
            {
                if (testSignature[i] != signature[i])
                {
                    throw new InvalidFileFormatException("Invalid file format.");
                }
            }

            classId = reader.ReadBytes(16);
            minorVersion = reader.ReadUInt16();
            majorVersion = reader.ReadUInt16();
            byteOrder = reader.ReadUInt16();
            sectorSize = reader.ReadUInt16();
            miniSectorSize = reader.ReadUInt16();
            reserved = reader.ReadUInt16();
            reserved1 = reader.ReadUInt32();
            directorySectorCount = reader.ReadUInt32();
            fatSectorCount = reader.ReadUInt32();
            firstDirectorySector = reader.ReadUInt32();
            transactionSignature = reader.ReadUInt32();
            miniStreamMaxSize = reader.ReadUInt32();
            firstMiniFatSector = reader.ReadUInt32();
            miniFatSectorCount = reader.ReadUInt32();
            firstDifatSector = reader.ReadUInt32();
            difatSectorCount = reader.ReadUInt32();

            for (int i = 0; i < 109; i++)
            {
                difat[i] = reader.ReadUInt32();
            }
        }

        internal byte[] ToBytes()
        {
            //For version 4 compound files, the header size (512 bytes) is less then the sector size (4096) bytes,
            //so the remaining part of the header (3584 bytes) must be filled with all zeros.
            byte[] buffer = new byte[SectorSize];

            MemoryStream memoryStream = new MemoryStream(buffer);

            using (memoryStream)
            {
                BinaryWriter writer = new BinaryWriter(memoryStream);

                writer.Write(signature);
                writer.Write(classId);
                writer.Write(minorVersion);
                writer.Write(majorVersion);
                writer.Write(byteOrder);
                writer.Write(sectorSize);
                writer.Write(miniSectorSize);
                writer.Write(reserved);
                writer.Write(reserved1);
                writer.Write(directorySectorCount);
                writer.Write(fatSectorCount);
                writer.Write(firstDirectorySector);
                writer.Write(transactionSignature);
                writer.Write(miniStreamMaxSize);
                writer.Write(firstMiniFatSector);
                writer.Write(miniFatSectorCount);
                writer.Write(firstDifatSector);
                writer.Write(difatSectorCount);

                for (int i = 0; i < 109; i++)
                {
                    writer.Write(difat[i]);
                }
            }

            return buffer;
        }

        #region Properties

        internal ushort MajorVersion
        {
            get
            {
                return majorVersion;
            }
            set
            {
                if (value == 3)
                {
                    majorVersion = value;
                    sectorSize = 0x0009;
                }
                else if (value == 4)
                {
                    majorVersion = value;
                    sectorSize = 0x000C;
                }
                else
                {
                    throw new ArgumentException("MajorVersion must be 3 or 4.");
                }
            }
        }

        internal ushort SectorSize
        {
            get
            {
                if (sectorSize == 9)
                {
                    return 512;
                }
                else 
                {
                    return 4096;
                }
            }
        }

        internal ushort MiniSectorSize
        {
            get
            {
                return 64;
            }
        }

        internal uint DirectorySectorCount
        {
            get
            {
                return directorySectorCount;
            }
            set
            {
                directorySectorCount = value;
            }
        }

        internal uint FatSectorCount
        {
            get
            {
                return fatSectorCount;
            }
            set
            {
                fatSectorCount = value;
            }
        }

        internal uint FirstDirectorySector
        {
            get
            {
                return firstDirectorySector;
            }
            set
            {
                firstDirectorySector = value;
            }
        }

        internal uint MiniStreamMaxSize
        {
            get
            {
                return miniStreamMaxSize;
            }
        }

        internal uint FirstMiniFatSector
        {
            get
            {
                return firstMiniFatSector;
            }
            set
            {
                firstMiniFatSector = value;
            }
        }

        internal uint MiniFatSectorCount
        {
            get
            {
                return miniFatSectorCount;
            }
            set
            {
                miniFatSectorCount = value;
            }
        }

        internal uint FirstDifatSector
        {
            get
            {
                return firstDifatSector;
            }
            set
            {
                firstDifatSector = value;
            }
        }

        internal uint DifatSectorCount
        {
            get
            {
                return difatSectorCount;
            }
            set
            {
                difatSectorCount = value;
            }
        }

        internal uint[] Difat
        {
            get
            {
                return difat;
            }
            set
            {
                difat = value;
            }
        }

        internal byte[] ClassId
        {
            get
            {
                return classId;
            }
            set
            {
                classId = value;
            }
        }

        #endregion
    }
}
