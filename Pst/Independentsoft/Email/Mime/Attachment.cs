using System;
using System.IO;
using System.Text;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Class Attachment.
    /// </summary>
    public class Attachment
    {
        private string name;
        private byte[] buffer;
        private ContentType contentType = new ContentType("application", "octet-stream");
        private string contentID;
        private string contentLocation;
        private string contentDescription;
        private ContentDisposition contentDisposition = new ContentDisposition(ContentDispositionType.Attachment);

        /// <summary>
        /// Initializes a new instance of the <see cref="Attachment"/> class.
        /// </summary>
        public Attachment()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attachment"/> class.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        public Attachment(string filePath) : this()
        {
            Open(filePath);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attachment"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public Attachment(Stream stream) : this()
        {
            Open(stream);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attachment"/> class.
        /// </summary>
        /// <param name="buffer">The buffer.</param>
        public Attachment(byte[] buffer) : this()
        {
            this.buffer = buffer;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attachment"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="name">The name.</param>
        public Attachment(Stream stream, string name) : this()
        {
            this.name = name;
            Open(stream);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attachment"/> class.
        /// </summary>
        /// <param name="buffer">The buffer.</param>
        /// <param name="name">The name.</param>
        public Attachment(byte[] buffer, string name) : this()
        {
            this.buffer = buffer;
            this.name = name;
        }

        public void Open(string filePath)
        {
            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);

            using (fileStream)
            {
                Open(fileStream);
            }

            FileInfo fileInfo = new FileInfo(filePath);
            this.name = fileInfo.Name;
        }

        private void Open(Stream stream)
        {
            byte[] tempBuffer = new byte[32768];
            int count = -1;

            MemoryStream memoryStream = new MemoryStream();

            using (memoryStream)
            {
                while ((count = stream.Read(tempBuffer, 0, tempBuffer.Length)) > 0)
                {
                    memoryStream.Write(tempBuffer, 0, count);
                }

                buffer = memoryStream.ToArray();
            }

            if (this.name != null)
            {
                Parameter nameParameter = new Parameter("filename", this.name);
                contentDisposition.Parameters.Add(nameParameter);
            }
        }

        /// <summary>
        /// Saves the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <exception cref="System.ArgumentNullException">stream</exception>
        public void Save(Stream stream)
        {
            if(stream == null)
            {
                throw new ArgumentNullException("stream");
            }

            if (buffer != null)
            {
                stream.Write(buffer, 0, buffer.Length);
            }
        }

        /// <summary>
        /// Saves this message to the specified file.
        /// </summary>
        /// <param name="filePath">File path.</param>
        public void Save(string filePath)
        {
            Save(filePath, false);
        }

        /// <summary>
        /// Saves this message to the specified file.
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
        /// Gets the stream.
        /// </summary>
        /// <returns>Stream.</returns>
        public Stream GetStream()
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

        /// <summary>
        /// Gets the bytes.
        /// </summary>
        /// <returns>System.Byte[][].</returns>
        public byte[] GetBytes()
        {
            return buffer;
        }

        /// <summary>
        /// Gets the name of the file.
        /// </summary>
        /// <returns>System.String.</returns>
        public string GetFileName()
        {
            if (name == null)
            {
                return "Attachment-" + Util.NextRandom(); 
            }
            else
            {
                string fileName = name;

                fileName = fileName.Replace(":", "_");
                fileName = fileName.Replace("*", "_");
                fileName = fileName.Replace("\\", "_");
                fileName = fileName.Replace("/", "_");
                fileName = fileName.Replace("?", "_");
                fileName = fileName.Replace("\"", "_");
                fileName = fileName.Replace("<", "_");
                fileName = fileName.Replace(">", "_");
                fileName = fileName.Replace("|", "_");
                fileName = fileName.Replace("", "_");
                fileName = fileName.Replace("\r", "_");
                fileName = fileName.Replace("\n", "_");
                fileName = fileName.Replace("\t", "_");

                if (fileName.Length > 128)
                {
                    int extensionIndex = fileName.LastIndexOf(".");

                    string extension = "";

                    if (extensionIndex > 124)
                    {
                        extension = fileName.Substring(extensionIndex);
                    }

                    fileName = fileName.Substring(0, 128) + extension;
                }

                return fileName;
            }
        }

        /// <summary>
        /// Returns a <see cref="System.String" /> that represents this instance.
        /// </summary>
        /// <returns>A <see cref="System.String" /> that represents this instance.</returns>
        public override string ToString()
        {
            return name;
        }

        #region Properties

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>The name.</value>
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

        /// <summary>
        /// Gets or sets the type of the content.
        /// </summary>
        /// <value>The type of the content.</value>
        public ContentType ContentType
        {
            get
            {
                return contentType;
            }
            set
            {
                contentType = value;
            }
        }

        /// <summary>
        /// Gets or sets the content identifier.
        /// </summary>
        /// <value>The content identifier.</value>
        public string ContentID
        {
            get
            {
                return contentID;
            }
            set
            {
                contentID = value;
            }
        }

        /// <summary>
        /// Gets or sets the content location.
        /// </summary>
        /// <value>The content location.</value>
        public string ContentLocation
        {
            get
            {
                return contentLocation;
            }
            set
            {
                contentLocation = value;
            }
        }

        /// <summary>
        /// Gets or sets the content description.
        /// </summary>
        /// <value>The content description.</value>
        public string ContentDescription
        {
            get
            {
                return contentDescription;
            }
            set
            {
                contentDescription = value;
            }
        }

        /// <summary>
        /// Gets or sets the content disposition.
        /// </summary>
        /// <value>The content disposition.</value>
        public ContentDisposition ContentDisposition
        {
            get
            {
                return contentDisposition;
            }
            set
            {
                contentDisposition = value;
            }
        }
        
        #endregion
    }
}
