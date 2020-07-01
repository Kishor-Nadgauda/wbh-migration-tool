using System;
using System.IO;

namespace Independentsoft.IO.StructuredStorage
{
    /// <summary>
    /// Represents a directory entry. 
    /// </summary>
    public abstract class DirectoryEntry : IComparable
    {
        internal string name;
        internal DirectoryEntryType type = DirectoryEntryType.Invalid;
        internal Color color = Color.Black;
        internal uint leftSiblingSid;
        internal uint rightSiblingSid;
        internal uint childSid;
        internal byte[] classId = new byte[16];
        internal uint userFlags;
        internal DateTime createdTime;
        internal DateTime lastModifiedTime;
        internal uint startSector;
        internal uint size;
        internal ushort propertyType = 0; //Reserved for the future use. Must be zero.
        internal byte[] buffer;
        internal DirectoryEntryList directoryEntries = new DirectoryEntryList();

        //during read
        internal DirectoryEntry parent;

        internal static DirectoryEntry Parse(BinaryReader reader)
        {
            byte[] nameValue = reader.ReadBytes(64);
            ushort nameLength = reader.ReadUInt16();

            string name = null;

            if (nameLength > 1)
            {
                name = System.Text.Encoding.Unicode.GetString(nameValue, 0, nameLength - 2);
            }
                        
            byte typeValue = reader.ReadByte();
            DirectoryEntryType type = EnumUtil.ParseDirectoryEntryType(typeValue);

            byte colorValue = reader.ReadByte();
            Color color = EnumUtil.ParseColor(colorValue);

            uint leftSiblingSid = reader.ReadUInt32();
            uint rightSiblingSid = reader.ReadUInt32();
            uint childSid = reader.ReadUInt32();
            byte[] classId = reader.ReadBytes(16);
            uint userFlags = reader.ReadUInt32();

            uint createdTimeLow = reader.ReadUInt32();
            long createdTimeHigh = reader.ReadUInt32();
            
            DateTime createdTime = new DateTime();

            if (createdTimeHigh > 0)
            {
                long ticks = createdTimeLow + (createdTimeHigh << 32);

                DateTime year1601 = new DateTime(1601, 1, 1);

                try
                {
                    createdTime = year1601.AddTicks(ticks).ToLocalTime();
                }
                catch (Exception) //ignore wrong dates
                {
                }
            }

            uint lastModifiedTimeLow = reader.ReadUInt32();
            long lastModifiedTimeHigh = reader.ReadUInt32();

            DateTime lastModifiedTime = new DateTime();

            if (lastModifiedTimeHigh > 0)
            {
                long ticks = lastModifiedTimeLow + (lastModifiedTimeHigh << 32);

                DateTime year1601 = new DateTime(1601, 1, 1);

                try
                {
                    lastModifiedTime = year1601.AddTicks(ticks).ToLocalTime();
                }
                catch (Exception) //ignore wrong dates
                {
                }
            }

            uint startSector = reader.ReadUInt32();
            uint size = reader.ReadUInt32();
            
            if (type == DirectoryEntryType.Root)
            {
                RootDirectoryEntry entry = new RootDirectoryEntry();
                entry.type = DirectoryEntryType.Root;
                entry.name = name;
                entry.color = color;
                entry.leftSiblingSid = leftSiblingSid;
                entry.rightSiblingSid = rightSiblingSid;
                entry.childSid = childSid;
                entry.classId = classId;
                entry.userFlags = userFlags;
                entry.createdTime = createdTime;
                entry.lastModifiedTime = lastModifiedTime;
                entry.startSector = startSector;
                entry.size = size;

                return entry;
            }
            else if (type == DirectoryEntryType.Stream)
            {
                Stream entry = new Stream();
                entry.type = DirectoryEntryType.Stream;
                entry.name = name;
                entry.color = color;
                entry.leftSiblingSid = leftSiblingSid;
                entry.rightSiblingSid = rightSiblingSid;
                entry.childSid = childSid;
                entry.classId = classId;
                entry.userFlags = userFlags;
                entry.createdTime = createdTime;
                entry.lastModifiedTime = lastModifiedTime;
                entry.startSector = startSector;
                entry.size = size;

                return entry;
            }
            else if (type == DirectoryEntryType.Storage)
            {
                Storage entry = new Storage();
                entry.type = DirectoryEntryType.Storage;
                entry.name = name;
                entry.color = color;
                entry.leftSiblingSid = leftSiblingSid;
                entry.rightSiblingSid = rightSiblingSid;
                entry.childSid = childSid;
                entry.classId = classId;
                entry.userFlags = userFlags;
                entry.createdTime = createdTime;
                entry.lastModifiedTime = lastModifiedTime;
                entry.startSector = startSector;
                entry.size = size;

                return entry;
            }
            else 
            {
                Storage entry = new Storage();
                entry.type = DirectoryEntryType.Invalid;
                entry.name = name;
                entry.color = color;
                entry.leftSiblingSid = leftSiblingSid;
                entry.rightSiblingSid = rightSiblingSid;
                entry.childSid = childSid;
                entry.classId = classId;
                entry.userFlags = userFlags;
                entry.createdTime = createdTime;
                entry.lastModifiedTime = lastModifiedTime;
                entry.startSector = startSector;
                entry.size = size;

                return entry;
            }
        }

        internal byte[] ToBytes()
        {
            byte[] buffer = new byte[128];

            MemoryStream memoryStream = new MemoryStream(buffer);

            using (memoryStream)
            {
                BinaryWriter writer = new BinaryWriter(memoryStream, System.Text.Encoding.Unicode);

                byte[] nameBuffer = new byte[64];
                byte[] unicodeNameBuffer = System.Text.Encoding.Unicode.GetBytes(name);

                for (int i = 0; i < unicodeNameBuffer.Length; i++)
                {
                    nameBuffer[i] = unicodeNameBuffer[i];
                }

                writer.Write(nameBuffer);
                writer.Write((ushort)((name.Length + 1) * 2));

                byte typeByte = EnumUtil.ParseDirectoryEntryType(type);
                writer.Write(typeByte);

                byte colorByte = EnumUtil.ParseColor(color);
                writer.Write(colorByte);

                writer.Write(leftSiblingSid);
                writer.Write(rightSiblingSid);
                writer.Write(childSid);
                writer.Write(classId);
                writer.Write(userFlags);

                if (createdTime.CompareTo(DateTime.MinValue) > 0)
                {
                    DateTime year1601 = new DateTime(1601, 1, 1);
                    TimeSpan timeSpan = createdTime.ToUniversalTime().Subtract(year1601);

                    long ticks = timeSpan.Ticks;

                    byte[] ticksBytes = BitConverter.GetBytes(ticks);

                    uint createdTimeLow = BitConverter.ToUInt32(ticksBytes, 0);
                    uint createdTimeHigh = BitConverter.ToUInt32(ticksBytes, 4);

                    writer.Write(createdTimeLow);
                    writer.Write(createdTimeHigh);
                }
                else
                {
                    writer.Write((uint)0);
                    writer.Write((uint)0);
                }

                if (lastModifiedTime.CompareTo(DateTime.MinValue) > 0)
                {
                    DateTime year1601 = new DateTime(1601, 1, 1);
                    TimeSpan timeSpan = lastModifiedTime.ToUniversalTime().Subtract(year1601);

                    long ticks = timeSpan.Ticks;

                    byte[] ticksBytes = BitConverter.GetBytes(ticks);

                    uint lastModifiedTimeLow = BitConverter.ToUInt32(ticksBytes, 0);
                    uint lastModifiedTimeHigh = BitConverter.ToUInt32(ticksBytes, 4);

                    writer.Write(lastModifiedTimeLow);
                    writer.Write(lastModifiedTimeHigh);
                }
                else
                {
                    writer.Write((uint)0);
                    writer.Write((uint)0);
                }

                writer.Write(startSector);
                writer.Write(size);
            }

            return buffer;
        }
        /// <summary>
        /// Compares this instance with the specified <see cref="DirectoryEntry"/> object and indicates whether this instance precedes, follows, or appears in the same position in the sort order as the specified DirectoryEntry.
        /// </summary>
        /// <param name="obj">A DirectoryEntry</param>
        /// <returns>A 32-bit signed integer that indicates whether this instance precedes, follows, or appears in the same position in the sort order as the value parameter.</returns>
        public int CompareTo(object obj)
        {
            if (obj is DirectoryEntry)
            {
                DirectoryEntry entry = (DirectoryEntry)obj;

                if (entry.Name.Length < name.Length)
                {
                    return 1;
                }
                else if (entry.Name.Length == name.Length)
                {
                    for (int i = 0; i < name.Length; i++)
                    {
                        int leftValue = (int)entry.Name[i];
                        int rightValue = (int)name[i];

                        if (leftValue < rightValue)
                        {
                            return 1;
                        }
                        else if (leftValue > rightValue)
                        {
                            return -1;
                        }
                    }

                    return 1;
                }
                else
                {
                    return -1;
                }
            }
            else
            {
                throw new ArgumentException("object is not a DirectoryEntry");
            }
        }

        #region Properties

        /// <summary>
        /// Gets or sets name.
        /// </summary>
        public string Name
        {
            get
            {
                return name;
            }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException("Name", "Name is null.");
                }

                if (value.Length > 31)
                {
                    throw new ArgumentException("Name", "Name must be less than 32 characters.");
                }

                name = value;
            }
        }

        /// <summary>
        /// Gets creation time.
        /// </summary>
        public DateTime CreatedTime
        {
            get
            {
                return createdTime;
            }
        }

        /// <summary>
        /// Gets last modified time.
        /// </summary>
        public DateTime LastModifiedTime
        {
            get
            {
                return lastModifiedTime;
            }
        }

        /// <summary>
        /// Gets size.
        /// </summary>
        public long Size
        {
            get
            {
                return size;
            }
        }

        /// <summary>
        /// Gets or sets class ID.
        /// </summary>
        public byte[] ClassId
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
