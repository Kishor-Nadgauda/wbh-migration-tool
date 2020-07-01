using System;
using System.IO;
using System.Collections.Generic;

namespace Independentsoft.IO.StructuredStorage
{
    /// <summary>
    /// Represents a file used to store data as virtual streams. 
    /// </summary>
    public class CompoundFile
    {
        private Header header = new Header();
        private RootDirectoryEntry root = new RootDirectoryEntry();

        /// <summary>
        /// Initializes a new instance of the CompoundFile class.  
        /// </summary>
        public CompoundFile()
        {
        }

        /// <summary>
        /// Initializes a new instance of the CompoundFile class based on the supplied file. 
        /// </summary>
        /// <param name="filePath">File path.</param>
        public CompoundFile(string filePath) : this()
        {
            Open(filePath);
        }

        /// <summary>
        /// Initializes a new instance of the CompoundFile class based on the supplied stream. 
        /// </summary>
        /// <param name="stream">A stream.</param>
        public CompoundFile(System.IO.Stream stream) : this()
        {
            Open(stream);
        }

        /// <summary>
        /// Opens compound file from the specified file.
        /// </summary>
        /// <param name="filePath">File path.</param>
        public void Open(string filePath)
        {
            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);

            using (fileStream)
            {
                Open(fileStream);
            }
        }

        /// <summary>
        /// Opens compound file from the specified stream.
        /// </summary>
        /// <param name="stream">A stream.</param>
        public void Open(System.IO.Stream stream)
        {
            OpenImplementation(stream);
        }

        private void OpenImplementation(System.IO.Stream stream)
        {
            if (stream.Length < 512)
            {
                throw new InvalidFileFormatException("Invalid file format.");
            }
            
            BinaryReader reader = new BinaryReader(stream, System.Text.Encoding.Unicode);
            reader.BaseStream.Position = 0;

            header = new Header(reader);

            uint[] difat = ParseDifat(reader);
            uint[] fat = ParseFat(reader, difat);
            uint[] miniFat = null;

            if (header.FirstMiniFatSector != 0xFFFFFFFE)
            {
                miniFat = ParseMiniFat(reader, fat);
            }

            IList<uint> directoryChain = new List<uint>();

            uint directorySector = header.FirstDirectorySector;
            directoryChain.Add(directorySector);

            while (true)
            {
                directorySector = fat[directorySector];

                if (directorySector != 0xFFFFFFFE)
                {
                    directoryChain.Add(directorySector);
                }
                else
                {
                    break;
                }
            }

            MemoryStream directoryMemoryStream = new MemoryStream();

            using (directoryMemoryStream)
            {
                BinaryReader directoryReader = new BinaryReader(directoryMemoryStream, System.Text.Encoding.Unicode);

                for (int i = 0; i < directoryChain.Count; i++)
                {
                    uint currentDirectoryChain = directoryChain[i];
                    reader.BaseStream.Position = currentDirectoryChain * header.SectorSize + header.SectorSize;

                    byte[] currentSector = reader.ReadBytes(header.SectorSize);
                    directoryMemoryStream.Write(currentSector, 0, currentSector.Length);
                }
                
                directoryReader.BaseStream.Position = 0;

                root = (RootDirectoryEntry)DirectoryEntry.Parse(directoryReader);

                IDictionary<uint, DirectoryEntry> directoryEntries = new Dictionary<uint, DirectoryEntry>();

                //read directory entries
                ////////////////////////////////////////////////
                directoryEntries.Add(0, root);

                if (root.childSid != 0xFFFFFFFF)
                {
                    directoryReader.BaseStream.Position = root.childSid * 128;
                    DirectoryEntry newEntry = DirectoryEntry.Parse(directoryReader);
                    
                    directoryEntries.Add(root.childSid, newEntry);
                    
                    root.directoryEntries.Add(newEntry);
                    newEntry.parent = root;

                    Stack<DirectoryEntry> leftStack = new Stack<DirectoryEntry>();
                    Stack<DirectoryEntry> rightStack = new Stack<DirectoryEntry>();
                    Stack<DirectoryEntry> childStack = new Stack<DirectoryEntry>();

                    leftStack.Push(newEntry);
                    rightStack.Push(newEntry);
                    childStack.Push(newEntry);

                    while (leftStack.Count > 0 || rightStack.Count > 0 || childStack.Count > 0)
                    {
                        if (leftStack.Count > 0)
                        {
                            DirectoryEntry currentEntry = (DirectoryEntry)leftStack.Pop();

                            if (currentEntry.leftSiblingSid != 0xFFFFFFFF && !directoryEntries.ContainsKey(currentEntry.leftSiblingSid))
                            {
                                directoryReader.BaseStream.Position = currentEntry.leftSiblingSid * 128;
                                newEntry = DirectoryEntry.Parse(directoryReader);

                                directoryEntries.Add(currentEntry.leftSiblingSid, newEntry);
                                
                                currentEntry.parent.directoryEntries.Add(newEntry);
                                newEntry.parent = currentEntry.parent;

                                leftStack.Push(newEntry);
                                rightStack.Push(newEntry);
                                childStack.Push(newEntry);

                                continue;
                            }
                        }

                        if (rightStack.Count > 0)
                        {
                            DirectoryEntry currentEntry = (DirectoryEntry)rightStack.Pop();

                            if (currentEntry.rightSiblingSid != 0xFFFFFFFF && !directoryEntries.ContainsKey(currentEntry.rightSiblingSid))
                            {
                                directoryReader.BaseStream.Position = currentEntry.rightSiblingSid * 128;
                                newEntry = DirectoryEntry.Parse(directoryReader);

                                directoryEntries.Add(currentEntry.rightSiblingSid, newEntry);

                                currentEntry.parent.directoryEntries.Add(newEntry);
                                newEntry.parent = currentEntry.parent;

                                leftStack.Push(newEntry);
                                rightStack.Push(newEntry);
                                childStack.Push(newEntry);

                                continue;
                            }
                        }

                        if (childStack.Count > 0)
                        {
                            DirectoryEntry currentEntry = (DirectoryEntry)childStack.Pop();

                            if (currentEntry.childSid != 0xFFFFFFFF && !directoryEntries.ContainsKey(currentEntry.childSid))
                            {
                                directoryReader.BaseStream.Position = currentEntry.childSid * 128;
                                newEntry = DirectoryEntry.Parse(directoryReader);

                                directoryEntries.Add(currentEntry.childSid, newEntry);

                                currentEntry.directoryEntries.Add(newEntry);
                                newEntry.parent = currentEntry;

                                leftStack.Push(newEntry);
                                rightStack.Push(newEntry);
                                childStack.Push(newEntry);
                            }
                        }
                    }
                }

                ////////////////////////////////////////////////

                IList<DirectoryEntry> directoryEntriesList = new List<DirectoryEntry>();

                foreach (DirectoryEntry entry in directoryEntries.Values)
                {
                    if (entry == root)
                    {
                        directoryEntriesList.Insert(0, entry);
                    }
                    else
                    {
                        directoryEntriesList.Add(entry);
                    }
                }

                MemoryStream miniStream = new MemoryStream();

                using (miniStream)
                {
                    BinaryReader miniStreamReader = new BinaryReader(miniStream);

                    for (int i = 0; i < directoryEntriesList.Count; i++)
                    {
                        DirectoryEntry entry = directoryEntriesList[i];

                        if (entry.type != DirectoryEntryType.Storage)
                        {
                            if (entry.type != DirectoryEntryType.Root && entry.Size > 0 && entry.Size < header.MiniStreamMaxSize)
                            {
                                IList<uint> streamChain = new List<uint>();

                                uint streamSector = entry.startSector;
                                streamChain.Add(streamSector);

                                while (true)
                                {
                                    streamSector = miniFat[streamSector];

                                    if (streamSector != 0xFFFFFFFC && streamSector != 0xFFFFFFFD && streamSector != 0xFFFFFFFE && streamSector != 0xFFFFFFFF && streamSector != miniFat[streamSector])
                                    {
                                        streamChain.Add(streamSector);
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }

                                MemoryStream entryBufferMemoryStream = new MemoryStream();

                                using (entryBufferMemoryStream)
                                {
                                    for (int j = 0; j < streamChain.Count; j++)
                                    {
                                        uint currentStreamChain = streamChain[j];
                                        miniStreamReader.BaseStream.Position = currentStreamChain * 64;

                                        byte[] currentSector = miniStreamReader.ReadBytes(64);
                                        entryBufferMemoryStream.Write(currentSector, 0, currentSector.Length);
                                    }

                                    entry.buffer = new byte[entry.size];

                                    if (entryBufferMemoryStream.Length < entry.buffer.Length)
                                    {
                                        entry.buffer = new byte[entryBufferMemoryStream.Length];
                                    }

                                    System.Array.Copy(entryBufferMemoryStream.ToArray(), 0, entry.buffer, 0, entry.buffer.Length);
                                }
                            }
                            else if (entry.Size > 0)
                            {
                                IList<uint> streamChain = new List<uint>();

                                uint streamSector = entry.startSector;
                                streamChain.Add(streamSector);

                                while (true)
                                {
                                    streamSector = fat[streamSector];

                                    if (streamSector != 0xFFFFFFFC && streamSector != 0xFFFFFFFD && streamSector != 0xFFFFFFFE && streamSector != 0xFFFFFFFF && streamSector != fat[streamSector])
                                    {
                                        streamChain.Add(streamSector);
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }

                                MemoryStream entryBufferMemoryStream = new MemoryStream();

                                using (entryBufferMemoryStream)
                                {
                                    for (int j = 0; j < streamChain.Count; j++)
                                    {
                                        uint currentStreamChain = streamChain[j];
                                        reader.BaseStream.Position = currentStreamChain * header.SectorSize + header.SectorSize;

                                        byte[] currentSector = reader.ReadBytes(header.SectorSize);
                                        entryBufferMemoryStream.Write(currentSector, 0, currentSector.Length);
                                    }

                                    entry.buffer = new byte[entry.Size];

                                    if (entryBufferMemoryStream.Length < entry.buffer.Length)
                                    {
                                        entry.buffer = new byte[entryBufferMemoryStream.Length];
                                    }

                                    System.Array.Copy(entryBufferMemoryStream.ToArray(), 0, entry.buffer, 0, entry.buffer.Length);
                                }

                                if (entry == root)
                                {
                                    if (root.buffer != null)
                                    {
                                        miniStream.Write(root.buffer, 0, root.buffer.Length);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets stream to read from this compound file.
        /// </summary>
        /// <returns></returns>
        public System.IO.Stream GetStream()
        {
            byte[] buffer = CreateCompoundFile();
            MemoryStream memoryStream = new MemoryStream(buffer);

            return memoryStream;
        }

        /// <summary>
        /// Gets buffer to read from this compound file.
        /// </summary>
        /// <returns></returns>
        public byte[] GetBytes()
        {
            return CreateCompoundFile();
        }

        /// <summary>
        /// Saves this compound file to the specified file.
        /// </summary>
        /// <param name="filePath">File path.</param>
        public void Save(string filePath)
        {
            Save(filePath, false);
        }

        /// <summary>
        /// Saves this compound file to the specified file.
        /// </summary>
        /// <param name="filePath">File path.</param>
        /// <param name="overwrite">True to overwrite existing file, otherwise false.</param>
        public void Save(string filePath, bool overwrite)
        {
            FileMode mode = FileMode.CreateNew;

            if (overwrite)
            {
                mode = FileMode.Create;
            }

            using (FileStream fileStream = new FileStream(filePath, mode, FileAccess.Write))
            {
                Save(fileStream);
            }
        }

        /// <summary>
        /// Saves this compound file to the specified stream.
        /// </summary>
        /// <param name="stream">A stream.</param>
        public void Save(System.IO.Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }

            byte[] buffer = CreateCompoundFile();
            stream.Write(buffer, 0, buffer.Length);
        }

        private byte[] CreateCompoundFile()
        {
            byte[] outputBuffer = null;

            IList<DirectoryEntry> allDirectoryEntries = new List<DirectoryEntry>();
            IList<uint> fat = new List<uint>();
            IList<uint> miniFat = new List<uint>();

            uint firstMiniFatSector = 0xFFFFFFFE;
            uint firstMiniStreamSector = 0xFFFFFFFE;
            uint miniFatSectorCount = 0;

            MemoryStream compoundFileMemoryStream = new MemoryStream();

            using (compoundFileMemoryStream)
            {
                BinaryWriter compoundFileWriter = new BinaryWriter(compoundFileMemoryStream, System.Text.Encoding.Unicode);

                root.color = Color.Red;
                root.type = DirectoryEntryType.Root;
                root.buffer = null;
                root.leftSiblingSid = 0xFFFFFFFF;
                root.rightSiblingSid = 0xFFFFFFFF;                
                root.createdTime = DateTime.MinValue;
                root.lastModifiedTime = DateTime.MinValue;
                root.size = 0;
                root.startSector = 0;

                allDirectoryEntries.Add(root);
                CreateDirectoryEntries(root, ref allDirectoryEntries);
                
                MemoryStream dataStream = new MemoryStream();
                MemoryStream miniStream = new MemoryStream();

                using (dataStream)
                {
                    using (miniStream)
                    {
                        BinaryWriter dataStreamWriter = new BinaryWriter(dataStream, System.Text.Encoding.Unicode);
                        BinaryWriter miniStreamWriter = new BinaryWriter(miniStream, System.Text.Encoding.Unicode);

                        for (int i = allDirectoryEntries.Count - 1; i >= 0; i--)
                        {
                            DirectoryEntry entry = allDirectoryEntries[i];

                            if (entry.buffer != null)
                            {
                                entry.size = (uint)entry.buffer.Length;
                            }
                            else
                            {
                                entry.size = 0;
                            }

                            //root entry
                            if (i == 0 && miniStream.Position > 0)
                            {
                                entry.buffer = new byte[miniStream.Position];
                                System.Array.Copy(miniStream.ToArray(), 0, entry.buffer, 0, (int)miniStream.Position);
                                entry.size = (uint)entry.buffer.Length;
                            }

                            if (i > 0 && entry.size > 0 && entry.size < header.MiniStreamMaxSize)
                            {
                                entry.startSector = (uint)miniFat.Count;

                                for (int offset = 0; offset < entry.buffer.Length; offset += header.MiniSectorSize)
                                {
                                    byte[] dataSector = new byte[header.MiniSectorSize];

                                    int copyLength = dataSector.Length;

                                    if (entry.buffer.Length < offset + header.MiniSectorSize)
                                    {
                                        copyLength = entry.buffer.Length - offset;
                                    }

                                    System.Array.Copy(entry.buffer, offset, dataSector, 0, copyLength);
                                    miniStreamWriter.Write(dataSector);

                                    if (offset + header.MiniSectorSize < entry.buffer.Length)
                                    {
                                        miniFat.Add((uint)miniFat.Count + 1);
                                    }
                                    else
                                    {
                                        miniFat.Add(0xFFFFFFFE);
                                    }
                                }
                            }
                            else if (i > 0 && entry.size > 0 && entry.size >= header.MiniStreamMaxSize)
                            {
                                entry.startSector = (uint)fat.Count;

                                for (int offset = 0; offset < entry.buffer.Length; offset += header.SectorSize)
                                {
                                    byte[] dataSector = new byte[header.SectorSize];

                                    int copyLength = dataSector.Length;

                                    if (entry.buffer.Length < offset + header.SectorSize)
                                    {
                                        copyLength = entry.buffer.Length - offset;
                                    }

                                    System.Array.Copy(entry.buffer, offset, dataSector, 0, copyLength);
                                    dataStreamWriter.Write(dataSector);

                                    if (offset + header.SectorSize < entry.buffer.Length)
                                    {
                                        fat.Add((uint)fat.Count + 1);
                                    }
                                    else
                                    {
                                        fat.Add(0xFFFFFFFE);
                                    }
                                }
                            }
                        }

                        int miniSectorCountInFatSector = header.SectorSize / header.MiniSectorSize;
                        int currentMiniSector = 0;
                        bool lastSector = false;

                        firstMiniStreamSector = (uint)fat.Count;

                        if (miniStream.Length > 0)
                        {
                            byte[] miniBuffer = miniStream.ToArray();
                            byte[] dataSector = new byte[header.SectorSize];

                            for (int offset = 0; offset < miniBuffer.Length; offset += header.MiniSectorSize)
                            {
                                int copyLength = header.MiniSectorSize;

                                if (miniBuffer.Length < offset + header.MiniSectorSize)
                                {
                                    copyLength = (int)(miniStream.Length - offset);
                                }

                                System.Array.Copy(miniBuffer, offset, dataSector, currentMiniSector * header.MiniSectorSize, copyLength);
                                currentMiniSector++;

                                if (offset + header.MiniSectorSize >= miniBuffer.Length)
                                {
                                    lastSector = true;
                                }

                                if (lastSector || currentMiniSector == miniSectorCountInFatSector)
                                {
                                    currentMiniSector = 0;
                                    dataStreamWriter.Write(dataSector);

                                    if (!lastSector)
                                    {
                                        fat.Add((uint)fat.Count + 1);
                                    }
                                    else
                                    {
                                        fat.Add(0xFFFFFFFE);
                                    }
                                }
                            }

                            int miniFatCountInFatSector = header.SectorSize / 4;
                            firstMiniFatSector = (uint)fat.Count;

                            for (int i = 0; i < miniFat.Count; i += miniFatCountInFatSector)
                            {
                                byte[] miniFatSector = new byte[header.SectorSize];

                                for (int j = 0; j < miniFatCountInFatSector; j++)
                                {
                                    if (i + j < miniFat.Count)
                                    {
                                        uint currentMiniFat = miniFat[i + j];

                                        byte[] currentMiniFatBuffer = BitConverter.GetBytes(currentMiniFat);
                                        System.Array.Copy(currentMiniFatBuffer, 0, miniFatSector, j * 4, 4);
                                    }
                                    else
                                    {
                                        byte[] currentMiniFatBuffer = BitConverter.GetBytes(0xFFFFFFFF);
                                        System.Array.Copy(currentMiniFatBuffer, 0, miniFatSector, j * 4, 4);
                                    }
                                }

                                dataStreamWriter.Write(miniFatSector);
                                miniFatSectorCount++;

                                if (i + miniFatCountInFatSector < miniFat.Count)
                                {
                                    fat.Add((uint)fat.Count + 1);
                                }
                                else
                                {
                                    fat.Add(0xFFFFFFFE);
                                }
                            }
                        }

                        compoundFileWriter.Write(dataStream.ToArray());
                    }
                }

                //Directory
                ///////////////////////////////////////////////////////////////////////////////////

                header.FirstDirectorySector = (uint)fat.Count;

                uint directorySectorCount = 0;

                int directoryEntriesInSector = header.SectorSize / 128;

                for (int i = 0; i < allDirectoryEntries.Count; i += directoryEntriesInSector)
                {
                    if (allDirectoryEntries[i] == root)
                    {
                        if (firstMiniStreamSector != 0xFFFFFFFE)
                        {
                            root.startSector = firstMiniStreamSector;
                        }
                    }

                    byte[] directorySector = new byte[header.SectorSize];

                    for (int j = 0; j < directoryEntriesInSector; j++)
                    {
                        if (i + j < allDirectoryEntries.Count)
                        {
                            DirectoryEntry entry = allDirectoryEntries[i + j];
                            System.Array.Copy(entry.ToBytes(), 0, directorySector, j * 128, 128);
                        }
                    }

                    compoundFileWriter.Write(directorySector);
                    directorySectorCount++;

                    if (i + directoryEntriesInSector < allDirectoryEntries.Count)
                    {
                        fat.Add((uint)fat.Count + 1);
                    }
                    else
                    {
                        fat.Add(0xFFFFFFFE);
                    }
                }

                if (header.MajorVersion == 4)
                {
                    header.DirectorySectorCount = directorySectorCount;
                }

                //DIFAT and FAT
                ///////////////////////////////////////////////////////////////////////////////////

                int fatSectorsPerSector = header.SectorSize / 4;
                int fatSectorCount = fat.Count / fatSectorsPerSector;

                if (fatSectorCount * fatSectorsPerSector < fat.Count)
                {
                    fatSectorCount = fatSectorCount + 1;
                }

                fatSectorCount = (fat.Count + fatSectorCount) / fatSectorsPerSector;

                if (fatSectorCount * fatSectorsPerSector < fat.Count + fatSectorCount)
                {
                    fatSectorCount = fatSectorCount + 1;
                }

                int difatSectorCount = (fatSectorCount - 109) / (fatSectorsPerSector - 1);

                if (difatSectorCount * fatSectorsPerSector < (fatSectorCount - 109))
                {
                    difatSectorCount = difatSectorCount + 1;
                }

                header.FatSectorCount = (uint)fatSectorCount;
              
                IList<uint> difatSectorList = new List<uint>();
                IList<uint> difatAddtionalSectorList = new List<uint>();

                for (int i = 0; i < fatSectorCount; i++)
                {
                    fat.Add(0xFFFFFFFD); //fat sectors
                    int position = fat.Count - 1;

                    if (i < 109)
                    {
                        difatSectorList.Add((uint)position);
                    }
                    else
                    {
                        difatAddtionalSectorList.Add((uint)position);
                    }
                }

                for (int i = 0; i < difatSectorCount; i++)
                {
                    fat.Add(0xFFFFFFFC); //difat sectors
                }

                for (int i = 0; i < fat.Count; i += fatSectorsPerSector)
                {
                    byte[] fatSector = new byte[header.SectorSize];

                    for (int j = 0; j < fatSectorsPerSector; j++)
                    {
                        if (i + j < fat.Count)
                        {
                            uint currentFat = (uint)fat[i + j];

                            byte[] currentFatBuffer = BitConverter.GetBytes(currentFat);
                            System.Array.Copy(currentFatBuffer, 0, fatSector, j * 4, 4);
                        }
                        else
                        {
                            byte[] currentFatBuffer = BitConverter.GetBytes(0xFFFFFFFF);
                            System.Array.Copy(currentFatBuffer, 0, fatSector, j * 4, 4);
                        }
                    }

                    compoundFileWriter.Write(fatSector);
                }

                if (difatSectorCount > 0)
                {
                    header.FirstDifatSector = (uint)(compoundFileWriter.BaseStream.Position / header.SectorSize);
                }
                else
                {
                    header.FirstDifatSector = 0xFFFFFFFE;
                }

                header.DifatSectorCount = (uint)difatSectorCount;

                for (int i = 0; i < difatSectorList.Count; i++)
                {
                    header.Difat[i] = difatSectorList[i];
                }

                for (int i = difatSectorList.Count; i < 109; i++)
                {
                    header.Difat[i] = 0xFFFFFFFF;
                }

                int nextDifatSector = 1;

                for (int i = 0; i < difatAddtionalSectorList.Count; i += fatSectorsPerSector - 1)
                {
                    byte[] fatSector = new byte[header.SectorSize];

                    for (int j = 0; j < fatSectorsPerSector - 1; j++)
                    {
                        if (i + j < difatAddtionalSectorList.Count)
                        {
                            uint currentFat = difatAddtionalSectorList[i + j];

                            byte[] currentFatBuffer = BitConverter.GetBytes(currentFat);
                            System.Array.Copy(currentFatBuffer, 0, fatSector, j * 4, 4);
                        }
                        else
                        {
                            byte[] currentFatBuffer = BitConverter.GetBytes(0xFFFFFFFF);
                            System.Array.Copy(currentFatBuffer, 0, fatSector, j * 4, 4);
                        }
                    }

                    if (i + (fatSectorsPerSector - 1) < difatAddtionalSectorList.Count)
                    {
                        byte[] nextDifatSectorBuffer = BitConverter.GetBytes(header.FirstDifatSector + nextDifatSector++);
                        System.Array.Copy(nextDifatSectorBuffer, 0, fatSector, (fatSectorsPerSector - 1) * 4, 4);
                    }
                    else
                    {
                        byte[] nextDifatSectorBuffer = BitConverter.GetBytes(0xFFFFFFFE);
                        System.Array.Copy(nextDifatSectorBuffer, 0, fatSector, (fatSectorsPerSector - 1) * 4, 4);
                    }

                    compoundFileWriter.Write(fatSector);
                }

                //Header
                ///////////////////////////////////////////////////////////////////////////////////

                header.FirstMiniFatSector = firstMiniFatSector;
                header.MiniFatSectorCount = miniFatSectorCount;

                byte[] headerBuffer = header.ToBytes();

                outputBuffer = new byte[compoundFileMemoryStream.Length + headerBuffer.Length];

                System.Array.Copy(headerBuffer, 0, outputBuffer, 0, headerBuffer.Length);
                System.Array.Copy(compoundFileMemoryStream.ToArray(), 0, outputBuffer, headerBuffer.Length, outputBuffer.Length - headerBuffer.Length);
            }

            return outputBuffer;
        }

        private void CreateDirectoryEntries(DirectoryEntry parent, ref IList<DirectoryEntry> allDirectoryEntries)
        {           
            if (parent.directoryEntries.Count > 0)
            {
                parent.directoryEntries.Sort();

                int middlePosition = parent.directoryEntries.Count / 2;

                DirectoryEntry middleEntry = parent.directoryEntries[middlePosition];

                if (parent.color == Color.Black)
                {
                    middleEntry.color = Color.Red;
                }
                else
                {
                    middleEntry.color = Color.Black;
                }

                middleEntry.createdTime = DateTime.Now;
                middleEntry.lastModifiedTime = middleEntry.createdTime;

                if (middleEntry.buffer != null)
                {
                    middleEntry.size = (uint)middleEntry.buffer.Length;
                }
                else
                {
                    middleEntry.size = 0;
                }

                middleEntry.startSector = 0;

                middleEntry.leftSiblingSid = 0xFFFFFFFF;
                middleEntry.rightSiblingSid = 0xFFFFFFFF;
                middleEntry.childSid = 0xFFFFFFFF;

                allDirectoryEntries.Add(middleEntry);
                parent.childSid = (uint)(allDirectoryEntries.Count - 1);

                DirectoryEntry previous = middleEntry;

                for (int l = middlePosition - 1; l >= 0; l--)
                {
                    DirectoryEntry leftEntry = parent.directoryEntries[l];

                    if (parent.color == Color.Black)
                    {
                        leftEntry.color = Color.Red;
                    }
                    else
                    {
                        leftEntry.color = Color.Black;
                    }

                    leftEntry.createdTime = DateTime.Now;
                    leftEntry.lastModifiedTime = leftEntry.createdTime;

                    if (leftEntry.buffer != null)
                    {
                        leftEntry.size = (uint)leftEntry.buffer.Length;
                    }
                    else
                    {
                        leftEntry.size = 0;
                    }

                    leftEntry.leftSiblingSid = 0xFFFFFFFF;
                    leftEntry.rightSiblingSid = 0xFFFFFFFF;
                    leftEntry.childSid = 0xFFFFFFFF;

                    allDirectoryEntries.Add(leftEntry);
                    previous.leftSiblingSid = (uint)(allDirectoryEntries.Count - 1);
                    previous = leftEntry;

                    if (leftEntry is Storage)
                    {
                        CreateDirectoryEntries(leftEntry, ref allDirectoryEntries);
                    }
                }

                previous = middleEntry;
                
                for (int r = middlePosition + 1; r < parent.directoryEntries.Count; r++)
                {
                    DirectoryEntry rightEntry = parent.directoryEntries[r];

                    if (parent.color == Color.Black)
                    {
                        rightEntry.color = Color.Red;
                    }
                    else
                    {
                        rightEntry.color = Color.Black;
                    }

                    rightEntry.createdTime = DateTime.Now;
                    rightEntry.lastModifiedTime = rightEntry.createdTime;

                    if (rightEntry.buffer != null)
                    {
                        rightEntry.size = (uint)rightEntry.buffer.Length;
                    }
                    else
                    {
                        rightEntry.size = 0;
                    }

                    rightEntry.leftSiblingSid = 0xFFFFFFFF;
                    rightEntry.rightSiblingSid = 0xFFFFFFFF;
                    rightEntry.childSid = 0xFFFFFFFF;

                    allDirectoryEntries.Add(rightEntry);
                    previous.rightSiblingSid = (uint)(allDirectoryEntries.Count - 1);
                    previous = rightEntry;

                    if (rightEntry is Storage)
                    {
                        CreateDirectoryEntries(rightEntry, ref allDirectoryEntries);
                    }
                }

                if (middleEntry is Storage)
                {
                    CreateDirectoryEntries(middleEntry, ref allDirectoryEntries);
                }
            }
        }

        private uint[] ParseFat(BinaryReader reader, uint[] difat)
        {
            int sectorCount = header.SectorSize / 4;

            uint[] fat = new uint[header.FatSectorCount * sectorCount];
            uint count = 0;

            for (int i = 0; i < difat.Length; i++)
            {
                reader.BaseStream.Position = difat[i] * header.SectorSize + header.SectorSize;

                for (int j = 0; j < sectorCount; j++)
                {
                    fat[count++] = reader.ReadUInt32();
                }
            }

            return fat;
        }

        private uint[] ParseMiniFat(BinaryReader reader, uint[] fat)
        {
            int sectorCount = header.SectorSize / 4;

            uint[] miniFat = new uint[header.MiniFatSectorCount * sectorCount];

            IList<uint> miniSectorChainList = new List<uint>();

            uint miniSector = header.FirstMiniFatSector;
            miniSectorChainList.Add(miniSector);

            while (true)
            {
                miniSector = fat[miniSector];

                if (miniSector != 0xFFFFFFFE)
                {
                    miniSectorChainList.Add(miniSector);
                }
                else
                {
                    break;
                }
            }

            uint[] miniSectorChain = new uint[miniSectorChainList.Count];

            for (int i = 0; i < miniSectorChain.Length; i++)
            {
                miniSectorChain[i] = UInt32.Parse(miniSectorChainList[i].ToString());
            }

            int count = 0;

            for (int j = 0; j < miniSectorChain.Length; j++)
            {
                reader.BaseStream.Position = miniSectorChain[j] * header.SectorSize + header.SectorSize;

                for (int k = 0; k < sectorCount; k++)
                {
                    miniFat[count++] = reader.ReadUInt32();
                }
            }

            return miniFat;
        }

        private uint[] ParseDifat(BinaryReader reader)
        {
            if (header.FatSectorCount <= 109)
            {
                uint[] difat = new uint[header.FatSectorCount];

                for (int i = 0; i < header.FatSectorCount; i++)
                {
                    difat[i] = header.Difat[i];
                }

                return difat;
            }
            else
            {
                int sectorCount = header.SectorSize / 4;
                uint[] sectorChain = new uint[sectorCount];

                uint[] difat = new uint[header.FatSectorCount];

                for (int i = 0; i < 109; i++)
                {
                    difat[i] = header.Difat[i];
                }

                reader.BaseStream.Position = header.FirstDifatSector * header.SectorSize + header.SectorSize;

                uint count = 109;

                while (true)
                {
                    for (int i = 0; i < sectorCount; i++)
                    {
                        sectorChain[i] = reader.ReadUInt32();
                    }

                    for (int i = 0; i < sectorCount - 1; i++)
                    {
                        if (sectorChain[i] != 0xFFFFFFFF && count < difat.Length)
                        {
                            difat[count++] = sectorChain[i];
                        }
                    }

                    if (sectorChain[sectorCount - 1] != 0xFFFFFFFE && count < difat.Length)
                    {
                        reader.BaseStream.Position = sectorChain[sectorCount - 1] * header.SectorSize + header.SectorSize;
                    }
                    else
                    {
                        break;
                    }
                }

                return difat;
            }
        }

        #region Properties

        /// <summary>
        /// Gets the root node.
        /// </summary>
        public RootDirectoryEntry Root
        {
            get
            {
                return root;
            }
        }

        /// <summary>
        /// Gets or sets major version. Allowed values are 3 (512 bytes sector size) or 4 (4096 bytes sector size).
        /// </summary>
        public int MajorVersion
        {
            get
            {
                return header.MajorVersion;
            }
            set
            {
                if (value == 3 || value == 4)
                {
                    header.MajorVersion = (ushort)value;
                }
                else
                {
                    throw new ArgumentException("MajorVersion must be 3 or 4.");
                }
            }
        }

        /// <summary>
        /// Gets FAT sector size.
        /// </summary>
        public int SectorSize
        {
            get
            {
                return header.SectorSize;
            }
        }

        /// <summary>
        /// Gets size of mini sectors.
        /// </summary>
        public int MiniSectorSize
        {
            get
            {
                return header.MiniSectorSize;
            }
        }

        /// <summary>
        /// Gets FAT sector count.
        /// </summary>
        public int FatSectorCount
        {
            get
            {
                return (int)header.FatSectorCount;
            }
        }

        /// <summary>
        /// Gets maximum size of mini streams.
        /// </summary>
        public long MiniStreamMaxSize
        {
            get
            {
                return header.MiniStreamMaxSize;
            }
        }

        /// <summary>
        /// Gets count of mini FAT sectors.
        /// </summary>
        public long MiniFatSectorCount
        {
            get
            {
                return header.MiniFatSectorCount;
            }
        }

        /// <summary>
        /// Gets or sets compound file class ID.
        /// </summary>
        public byte[] ClassId
        {
            get
            {
                return header.ClassId;
            }
            set
            {
                if (value != null)
                {
                    header.ClassId = value;
                }
            }
        }

        #endregion
    }
}
