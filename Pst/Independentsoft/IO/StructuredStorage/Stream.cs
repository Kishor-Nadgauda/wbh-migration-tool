using System;
using System.IO;

namespace Independentsoft.IO.StructuredStorage
{
    /// <summary>
    /// Represents a virtual stream to store data.
    /// </summary>
    /// <remarks>
    /// Stream is analogous to a file system file. The parent object of a stream object must be a <see cref="Storage"/> object or the <see cref="RootDirectoryEntry"/>.
    /// </remarks>
    public class Stream : DirectoryEntry
    {
        /// <summary>
        /// Initializes a new instance of the Stream class.  
        /// </summary>
        public Stream()
        {
            this.type = DirectoryEntryType.Stream;
        }

        /// <summary>
        /// Initializes a new instance of the Stream class and load data from the specified file.
        /// </summary>
        /// <param name="filePath">File path.</param>
        public Stream(string filePath) : this()
        {
            Load(filePath);
        }

        /// <summary>
        /// Initializes a new instance of the Stream class and load data from the specified <see cref="System.IO.Stream"/>.
        /// </summary>
        /// <param name="name">Stream name.</param>
        /// <param name="stream">A stream.</param>
        public Stream(string name, System.IO.Stream stream) : this()
        {
            Load(name, stream);
        }

        /// <summary>
        /// Initializes a new instance of the Stream class and load data from the specified buffer.
        /// </summary>
        /// <param name="name">Stream name.</param>
        /// <param name="buffer">Data buffer.</param>
        public Stream(string name, byte[] buffer) : this()
        {
            this.name = name;
            this.buffer = buffer;
        }

        /// <summary>
        /// Loads data to this stream from the specified file.
        /// </summary>
        /// <param name="filePath">File path.</param>
        public void Load(string filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);
            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);

            using (fileStream)
            {
                Load(fileInfo.Name, fileStream);
            }
        }

        /// <summary>
        /// Loads data to this stream from the specified <see cref="System.IO.Stream"/>.
        /// </summary>
        /// <param name="name">Stream name.</param>
        /// <param name="stream">A stream.</param>
        public void Load(string name, System.IO.Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }

            this.name = name;

            buffer = new byte[stream.Length];

            stream.Read(buffer, 0, buffer.Length);
        }

        /// <summary>
        /// Loads data to this stream from the specified buffer.
        /// </summary>
        /// <param name="name">Stream name.</param>
        /// <param name="buffer">Data buffer.</param>
        public void Load(string name, byte[] buffer)
        {
            this.name = name;
            this.buffer = buffer;
        }

        /// <summary>
        /// Saves data from this stream to to the specified file.
        /// </summary>
        /// <param name="filePath">File path.</param>
        public void Save(string filePath)
        {
            if (buffer != null && buffer.Length > 0)
            {
                FileStream fileStream = new FileStream(filePath, FileMode.Create);

                using (fileStream)
                {
                    fileStream.Write(buffer, 0, buffer.Length);
                }
            }
        }

        /// <summary>
        /// Saves data from this stream to the specified <see cref="System.IO.Stream"/>.
        /// </summary>
        /// <param name="stream">A stream.</param>
        public void Save(System.IO.Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }

            if (buffer != null && buffer.Length > 0)
            {
                stream.Write(buffer, 0, buffer.Length);
            }
        }

        /// <summary>
        /// Gets <see cref="System.IO.Stream"/> to read data from this stream. 
        /// </summary>
        /// <returns></returns>
        public System.IO.Stream GetStream()
        {
            if (buffer != null)
            {
                MemoryStream memoryStream = new MemoryStream(buffer);
                return memoryStream;
            }
            else
            {
                return null;
            }
        }

        #region Properties

        /// <summary>
        /// Gets or sets streams data.
        /// </summary>
        public byte[] Buffer
        {
            get
            {
                return buffer;
            }
            set
            {
                buffer = value;
            }
        }

        #endregion
    }
}
