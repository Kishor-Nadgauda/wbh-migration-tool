
using System;
using System.Collections.Generic;

namespace Independentsoft.IO.StructuredStorage
{
    /// <summary>
    /// Description of DirectoryEntryList.
    /// </summary>
	public class DirectoryEntryList : List<DirectoryEntry>
	{
        /// <summary>
        /// Gets the <see cref="DirectoryEntry"/> with the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns>DirectoryEntry.</returns>
        public DirectoryEntry this[string name]
        {
            get
            {
                for (int i = 0; i < this.Count; i++)
                {
                    DirectoryEntry currentEntry = this[i];

                    if (currentEntry.Name == name)
                    {
                        return currentEntry;
                    }
                }
                
                return null;
            }
        }
	}
}
