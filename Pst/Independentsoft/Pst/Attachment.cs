using System;
using System.IO;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class Attachment.
    /// </summary>
    public class Attachment
    {
        internal byte[] additionalInfo;
        internal string contentBase;
        internal string contentId;
        internal string contentLocation;
        internal string contentDisposition;
        internal byte[] data;
        internal byte[] dataObject;
        internal byte[] encoding;
        internal byte[] recordKey;
        internal string extension;
        internal string fileName;
        internal AttachmentFlags flags = AttachmentFlags.None;
        internal string longFileName;
        internal string longPathName;
        internal AttachmentMethod method = AttachmentMethod.AttachByValue;
        internal long mimeSequence;
        internal string mimeTag;
        internal string pathName;
        internal byte[] rendering;
        internal long renderingPosition;
        internal long size;
        internal byte[] tag;
        internal string transportName;
        internal string displayName;
        internal Item embeddedItem;
        internal ObjectType objectType = ObjectType.None;
        internal bool isHidden;
        internal bool isContactPhoto;
        internal DateTime creationTime;
        internal DateTime lastModificationTime;

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
        public Attachment(string filePath)
        {
            Load(filePath);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attachment"/> class.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="stream">The stream.</param>
        public Attachment(string fileName, Stream stream)
        {
            Load(fileName, stream);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attachment"/> class.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="buffer">The buffer.</param>
        public Attachment(string fileName, byte[] buffer)
        {
            Load(fileName, buffer);
        }

        private void Load(string filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);

            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);

            using (fileStream)
            {
                Load(fileInfo.Name, fileStream);
            }
        }

        private void Load(string fileName, Stream stream)
        {
            this.fileName = fileName;

            MemoryStream memoryStream = new MemoryStream();
            byte[] buffer = new byte[8192];
            int len = 0;

            using (memoryStream)
            {
                while ((len = stream.Read(buffer, 0, buffer.Length)) > 0)
                {
                    memoryStream.Write(buffer, 0, len);
                }

                data = memoryStream.ToArray();
            }
        }

        private void Load(string fileName, byte[] buffer)
        {
            this.fileName = fileName;
            this.data = buffer;
        }

        /// <summary>
        /// Saves the specified file path.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        public void Save(string filePath)
        {
            Save(filePath, false);
        }

        /// <summary>
        /// Saves the specified file path.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        public void Save(string filePath, bool overwrite)
        {
            FileMode mode = FileMode.CreateNew;

            if (overwrite)
            {
                mode = FileMode.Create;
            }

            FileStream fileStream = new FileStream(filePath, mode, FileAccess.Write);

            using (fileStream)
            {
                if (data != null && data.Length > 0)
                {
                    fileStream.Write(data, 0, data.Length);
                }
                else if (dataObject != null && dataObject.Length > 0)
                {
                    fileStream.Write(dataObject, 0, dataObject.Length);
                }
            }
        }

        /// <summary>
        /// Saves the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public void Save(Stream stream)
        {
            if (data != null && data.Length > 0)
            {
                stream.Write(data, 0, data.Length);
            }
            else if (dataObject != null && dataObject.Length > 0)
            {
                stream.Write(dataObject, 0, dataObject.Length);
            }
        }

        /// <summary>
        /// Gets the bytes.
        /// </summary>
        /// <returns>System.Byte[][].</returns>
        public byte[] GetBytes()
        {
            if (data != null)
            {
                return data;
            }
            else if (dataObject != null)
            {
                return dataObject;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Gets the buffer.
        /// </summary>
        /// <returns>System.Byte[][].</returns>
        public byte[] GetBuffer()
        {
            if (data != null)
            {
                return data;
            }
            else if (dataObject != null)
            {
                return dataObject;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Gets the stream.
        /// </summary>
        /// <returns>Stream.</returns>
        public Stream GetStream()
        {
            if (data != null)
            {
                MemoryStream memoryStream = new MemoryStream(data);
                return memoryStream;
            }
            else if (dataObject != null)
            {
                MemoryStream memoryStream = new MemoryStream(dataObject);
                return memoryStream;
            }
            else
            {
                return null;
            }
        }

        #region Properties

        /// <summary>
        /// Provides file type information for a non-Windows attachment.
        /// </summary>
        /// <value>The additional information.</value>
        public byte[] AdditionalInfo
        {
            get
            {
                return additionalInfo;
            }
            set
            {
                additionalInfo = value;
            }
        }

        /// <summary>
        /// Contains the content base header of a MIME message attachment.
        /// </summary>
        /// <value>The content base.</value>
        public string ContentBase
        {
            get
            {
                return contentBase;
            }
            set
            {
                contentBase = value;
            }
        }

        /// <summary>
        /// Contains the content identification header of a MIME message attachment.
        /// </summary>
        /// <value>The content identifier.</value>
        public string ContentId
        {
            get
            {
                return contentId;
            }
            set
            {
                contentId = value;
            }
        }

        /// <summary>
        /// Contains the content location header of a MIME message attachment.
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
        /// Contains the content disposition header of a MIME message attachment.
        /// </summary>
        /// <value>The content disposition.</value>
        public string ContentDisposition
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

        /// <summary>
        /// Contains binary attachment data.
        /// </summary>
        /// <value>The data.</value>
        public byte[] Data
        {
            get
            {
                return data;
            }
            set
            {
                data = value;
            }
        }

        /// <summary>
        /// Contains attachment's data as embedded object.
        /// </summary>
        /// <value>The data object.</value>
        public byte[] DataObject
        {
            get
            {
                return dataObject;
            }
            set
            {
                dataObject = value;
            }
        }

        /// <summary>
        /// Contains the encoding for an attachment.
        /// </summary>
        /// <value>The encoding.</value>
        public byte[] Encoding
        {
            get
            {
                return encoding;
            }
            set
            {
                encoding = value;
            }
        }

        /// <summary>
        /// Contains the record key for an attachment.
        /// </summary>
        /// <value>The record key.</value>
        public byte[] RecordKey
        {
            get
            {
                return recordKey;
            }
            set
            {
                recordKey = value;
            }
        }

        /// <summary>
        /// Contains a file name extension that indicates the document type of an attachment.
        /// </summary>
        /// <value>The extension.</value>
        public string Extension
        {
            get
            {
                return extension;
            }
            set
            {
                extension = value;
            }
        }

        /// <summary>
        /// Contains an attachment's base file name and extension, excluding path.
        /// </summary>
        /// <value>The name of the file.</value>
        public string FileName
        {
            get
            {
                return fileName;
            }
            set
            {
                fileName = value;
            }
        }

        /// <summary>
        /// Contains flags for an attachment.
        /// </summary>
        /// <value>The flags.</value>
        public AttachmentFlags Flags
        {
            get
            {
                return flags;
            }
            set
            {
                flags = value;
            }
        }

        /// <summary>
        /// Contains an attachment's long filename and extension, excluding path.
        /// </summary>
        /// <value>The long name of the file.</value>
        public string LongFileName
        {
            get
            {
                return longFileName;
            }
            set
            {
                longFileName = value;
            }
        }

        /// <summary>
        /// Contains an attachment's fully-qualified long path and filename.
        /// </summary>
        /// <value>The long name of the path.</value>
        public string LongPathName
        {
            get
            {
                return longPathName;
            }
            set
            {
                longPathName = value;
            }
        }

        /// <summary>
        /// Contains a MAPI-defined constant representing the way the contents of an attachment can be accessed.
        /// </summary>
        /// <value>The method.</value>
        public AttachmentMethod Method
        {
            get
            {
                return method;
            }
            set
            {
                method = value;
            }
        }

        /// <summary>
        /// Contains the MIME sequence number of a MIME message attachment.
        /// </summary>
        /// <value>The MIME sequence.</value>
        public long MimeSequence
        {
            get
            {
                return mimeSequence;
            }
            set
            {
                mimeSequence = value;
            }
        }

        /// <summary>
        /// Contains formatting information about a MIME attachment.
        /// </summary>
        /// <value>The MIME tag.</value>
        public string MimeTag
        {
            get
            {
                return mimeTag;
            }
            set
            {
                mimeTag = value;
            }
        }

        /// <summary>
        /// Contains an attachment's fully-qualified path and filename.
        /// </summary>
        /// <value>The name of the path.</value>
        public string PathName
        {
            get
            {
                return pathName;
            }
            set
            {
                pathName = value;
            }
        }

        /// <summary>
        /// Contains a Microsoft Windows metafile with rendering information for an attachment.
        /// </summary>
        /// <value>The rendering.</value>
        public byte[] Rendering
        {
            get
            {
                return rendering;
            }
            set
            {
                rendering = value;
            }
        }

        /// <summary>
        /// Contains rendering position index.
        /// </summary>
        /// <value>The rendering position.</value>
        public long RenderingPosition
        {
            get
            {
                return renderingPosition;
            }
            set
            {
                renderingPosition = value;
            }
        }

        /// <summary>
        /// Contains attachment's size in bytes.
        /// </summary>
        /// <value>The size.</value>
        public long Size
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

        /// <summary>
        /// Contains an object identifier specifying the application that supplied an attachment.
        /// </summary>
        /// <value>The tag.</value>
        public byte[] Tag
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

        /// <summary>
        /// Contains the name of an attachment file modified so that it can be associated with TNEF messages.
        /// </summary>
        /// <value>The name of the transport.</value>
        public string TransportName
        {
            get
            {
                return transportName;
            }
            set
            {
                transportName = value;
            }
        }

        /// <summary>
        /// Contains the display name of the attachment.
        /// </summary>
        /// <value>The display name.</value>
        public string DisplayName
        {
            get
            {
                return displayName;
            }
            set
            {
                displayName = value;
            }
        }

        /// <summary>
        /// Contains <see cref="Item" /> object if the attachment is an emabedded item.
        /// </summary>
        /// <value>The embedded item.</value>
        public Item EmbeddedItem
        {
            get
            {
                return embeddedItem;
            }
            set
            {
                if (value != null)
                {
                    embeddedItem = value;
                    embeddedItem.isEmbedded = true;
                }
            }
        }

        /// <summary>
        /// Contains the type of the attachment.
        /// </summary>
        /// <value>The type of the object.</value>
        public ObjectType ObjectType
        {
            get
            {
                return objectType;
            }
            set
            {
                objectType = value;
            }
        }

        /// <summary>
        /// Indicates whether an attachment is hidden from the end user.
        /// </summary>
        /// <value><c>true</c> if this instance is hidden; otherwise, <c>false</c>.</value>
        public bool IsHidden
        {
            get
            {
                return isHidden;
            }
            set
            {
                isHidden = value;
            }
        }

        /// <summary>
        /// Indicates whether this attachment is a contact photo.
        /// </summary>
        /// <value><c>true</c> if this instance is contact photo; otherwise, <c>false</c>.</value>
        public bool IsContactPhoto
        {
            get
            {
                return isContactPhoto;
            }
            set
            {
                isContactPhoto = value;
            }
        }

        /// <summary>
        /// Gets or sets the creation time.
        /// </summary>
        /// <value>The creation time.</value>
        public DateTime CreationTime
        {
            get
            {
                return creationTime;
            }
            set
            {
                creationTime = value;
            }
        }

        /// <summary>
        /// Gets or sets the last modification time.
        /// </summary>
        /// <value>The last modification time.</value>
        public DateTime LastModificationTime
        {
            get
            {
                return lastModificationTime;
            }
            set
            {
                lastModificationTime = value;
            }
        }

        #endregion
    }
}
