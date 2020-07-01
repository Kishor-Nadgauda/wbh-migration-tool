using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Represents the way the contents of an attachment can be accessed.
    /// </summary>
    public enum AttachmentMethod
    {
        /// <summary>
        /// The attachment has just been created.
        /// </summary>
        NoAttachment,
        
        /// <summary>
        /// The <see cref="Attachment.DataObject"/> property contains the attachment data.
        /// </summary>
        AttachByValue,
        
        /// <summary>
        /// The <see cref="Attachment.PathName"/> or the <see cref="Attachment.LongPathName"/> property contains a fully-qualified path identifying the attachment to recipients with access to a common file server.
        /// </summary>
        AttachByReference,
        
        /// <summary>
        /// The <see cref="Attachment.PathName"/> or the <see cref="Attachment.LongPathName"/> property contains a fully-qualified path identifying the attachment.
        /// </summary>
        AttachByReferenceResolve,
        
        /// <summary>
        /// The <see cref="Attachment.PathName"/> or the <see cref="Attachment.LongPathName"/> property contains a fully-qualified path identifying the attachment.
        /// </summary>
        AttachByReferenceOnly,
        
        /// <summary>
        /// The <see cref="Attachment.DataObject"/> property contains an embedded <see cref="Message"/> object.
        /// </summary>
        EmbeddedMessage,
        
        /// <summary>
        /// The attachment is an embedded OLE object
        /// </summary>
        Ole,

        /// <summary>
        /// None
        /// </summary>
        None
    }
}
