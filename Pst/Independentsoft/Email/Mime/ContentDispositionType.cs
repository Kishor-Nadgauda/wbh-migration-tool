using System;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Summary description for ContentDispositionType.
    /// </summary>
    public enum ContentDispositionType
    {
        /// <summary>
        /// Specifies that the attachment is to be displayed as a file attached to the e-mail message.
        /// </summary>
        Attachment,

        /// <summary>
        /// The attachment is to be displayed as part of the e-mail message body.
        /// </summary>
        Inline
    }
}
