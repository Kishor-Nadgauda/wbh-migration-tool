using System;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class MessageStore.
    /// </summary>
    public class MessageStore
    {
        private string displayName;
        private string comment;
        private byte[] recordKey;
        private uint passwordCheckSum;
        private Table table;

        internal MessageStore(Table table)
        {
            Parse(table);
        }

        private void Parse(Table table)
        {
            this.table = table;

            if (table.Entries[MapiPropertyTag.PR_DISPLAY_NAME] != null)
            {
                displayName = table.Entries[MapiPropertyTag.PR_DISPLAY_NAME].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_COMMENT] != null)
            {
                comment = table.Entries[MapiPropertyTag.PR_COMMENT].GetStringValue();
            }

            if (table.Entries[MapiPropertyTag.PR_RECORD_KEY] != null)
            {
                recordKey = table.Entries[MapiPropertyTag.PR_RECORD_KEY].GetBinaryValue();
            }

            if (table.Entries[MapiPropertyTag.PR_PST_PASSWORD] != null)
            {
                passwordCheckSum = BitConverter.ToUInt32(table.Entries[MapiPropertyTag.PR_PST_PASSWORD].GetBinaryValue(), 0);
            }
        }

        #region Properties

        /// <summary>
        /// Gets the display name.
        /// </summary>
        /// <value>The display name.</value>
        public string DisplayName
        {
            get
            {
                return displayName;
            }
        }

        /// <summary>
        /// Gets the comment.
        /// </summary>
        /// <value>The comment.</value>
        public string Comment
        {
            get
            {
                return comment;
            }
        }

        /// <summary>
        /// Gets the record key.
        /// </summary>
        /// <value>The record key.</value>
        public byte[] RecordKey
        {
            get
            {
                return recordKey;
            }
        }

        /// <summary>
        /// Gets the password check sum.
        /// </summary>
        /// <value>The password check sum.</value>
        public long PasswordCheckSum
        {
            get
            {
                return passwordCheckSum;
            }
        }

        /// <summary>
        /// Gets the table.
        /// </summary>
        /// <value>The table.</value>
        public Table Table
        {
            get
            {
                return table;
            }
        }

        #endregion
    }
}
