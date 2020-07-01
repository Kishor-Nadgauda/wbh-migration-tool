using System;
using System.IO;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class Table.
    /// </summary>
    public abstract class Table
    {
        private TableEntryList entries = new TableEntryList();
        private TableEntryList[] entriesArray = new TableEntryList[0];
        private IDictionary<uint, SLEntry> slEntriesTable = new Dictionary<uint, SLEntry>();

        internal TableEntry GetEntry(ushort tag, PropertyType type)
        {
            for (int i = 0; i < entries.Count; i++)
            {
                if (entries[i].PropertyTag.Tag == tag && entries[i].PropertyTag.Type == type)
                {
                    return entries[i];
                }
            }

            return null;
        }

        internal TableEntry GetEntry(int tag)
        {
            for (int i = 0; i < entries.Count; i++)
            {
                if (entries[i].PropertyTag.Tag == tag)
                {
                    return entries[i];
                }
            }

            return null;
        }

        #region Properties

        /// <summary>
        /// Gets the entries.
        /// </summary>
        /// <value>The entries.</value>
        public TableEntryList Entries
        {
            get
            {
                return entries;
            }
        }

        /// <summary>
        /// Gets or sets the entries array.
        /// </summary>
        /// <value>The entries array.</value>
        public TableEntryList[] EntriesArray
        {
            get
            {
                return entriesArray;
            }
            set
            {
                entriesArray = value;
            }
        }

        internal IDictionary<uint, SLEntry> SLEntriesTable
        {
            get
            {
                return slEntriesTable;
            }
        }

        #endregion
    }
}