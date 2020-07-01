using System;
using System.IO;

namespace Independentsoft.IO.StructuredStorage
{
    /// <summary>
    /// Represents a root node.
    /// </summary>
    /// <remarks>
    /// RootDirectoryEntry object in a compound file that must be accessed before any other <see cref="Storage"/> objects and <see cref="Stream"/> objects are referenced. It is the uppermost parent object in the storage object and stream object hierarchy.
    /// </remarks>
    public class RootDirectoryEntry : DirectoryEntry
    {
        internal RootDirectoryEntry()
        {
            this.name = "Root Entry";
            this.type = DirectoryEntryType.Root;
        }

        #region Properties

        /// <summary>
        /// Gets root's name.
        /// </summary>
        public new string Name
        {
            get
            {
                return name;
            }
        }

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
