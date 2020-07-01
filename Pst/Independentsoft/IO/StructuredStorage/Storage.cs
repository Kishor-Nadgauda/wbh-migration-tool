using System;

namespace Independentsoft.IO.StructuredStorage
{
    /// <summary>
    /// Contains collection of <see cref="DirectoryEntry"/>.
    /// </summary>
    /// <remarks>
    /// Storage is analogous to a file system directory. The parent object of a storage object must be another storage object or the <see cref="RootDirectoryEntry"/>.
    /// </remarks>
    public class Storage : DirectoryEntry
    {
        /// <summary>
        ///  Initializes a new instance of the Storage class.  
        /// </summary>
        public Storage()
        {
            this.type = DirectoryEntryType.Storage;
        }

        /// <summary>
        ///  Initializes a new instance of the Storage class.  
        /// </summary>
        /// <param name="name">Storage name.</param>
        public Storage(string name) : this()
        {
            this.name = name;
        }

        #region Properties

        /// <summary>
        /// Gets collection of <see cref="DirectoryEntry"/>.
        /// </summary>
        public DirectoryEntryList DirectoryEntries
        {
            get
            {
                return directoryEntries;
            }
        }

        #endregion
    }
}
