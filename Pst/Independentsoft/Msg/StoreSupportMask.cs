using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Contains a value that client applications should query to determine the characteristics of a message store.
    /// </summary>
    public enum StoreSupportMask
    {
        /// <summary>
        /// The message store supports properties containing ANSI (8-bit) characters.
        /// </summary>
        Ansi,

        /// <summary>
        /// The message store supports attachments (OLE or non-OLE) to messages.
        /// </summary>
        Attachments,

        /// <summary>
        /// The message store supports categorized views of tables.
        /// </summary>
        Categorize,

        /// <summary>
        /// The message store supports creation of new messages.
        /// </summary>
        Create,

        /// <summary>
        /// Entry identifiers for the objects in the message store are unique, that is, never reused during the life of the store.
        /// </summary>
        EntryIdUnique,

        /// <summary>
        /// The message store supports HTML messages, stored in the <see cref="Message.BodyHtml"/> property.
        /// </summary>
        Html,

        /// <summary>
        /// In a wrapped PST store, indicates that when a new message arrives at the store, the store does rules and spam filter processing on the message separately.
        /// </summary>
        ItemProc,

        /// <summary>
        /// This flag is reserved and should not be used.
        /// </summary>
        LocalStore,

        /// <summary>
        /// The message store supports modification of its existing messages.
        /// </summary>
        Modify,

        /// <summary>
        /// The message store supports multivalued properties, guarantees the stability of value order in a multivalued property throughout a save operation, and supports instantiation of multivalued properties in tables.
        /// </summary>
        MultiValueProperties,

        /// <summary>
        /// The message store supports notifications.
        /// </summary>
        Notify,

        /// <summary>
        /// The message store supports OLE attachments. 
        /// </summary>
        Ole,

        /// <summary>
        /// The folders in this store are public (multi-user), not private (possibly multi-instance but not multi-user).
        /// </summary>
        PublicFolders,

        /// <summary>
        /// The MAPI Protocol Handler will not crawl the store, and the store is responsible to push any changes through notifications to the indexer to have messages indexed.
        /// </summary>
        Pusher,

        /// <summary>
        /// All interfaces for the message store have a read-only access level.
        /// </summary>
        ReadOnly,

        /// <summary>
        /// The message store supports restrictions.
        /// </summary>
        Restrictions,

        /// <summary>
        /// The message store supports Rich Text Format (RTF) messages, usually compressed.
        /// </summary>
        Rtf,

        /// <summary>
        /// The message store supports search-results folders.
        /// </summary>
        Search,

        /// <summary>
        /// The message store supports sorting views of tables.
        /// </summary>
        Sort,

        /// <summary>
        /// The message store supports marking a message for submission.
        /// </summary>
        Submit,

        /// <summary>
        /// The message store supports storage of RTF messages in uncompressed form.
        /// </summary>
        UncompressedRtf,

        /// <summary>
        /// The message store supports properties containing Unicode characters.
        /// </summary>
        Unicode
    }
}
