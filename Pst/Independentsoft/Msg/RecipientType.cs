using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Represents the recipient type for a message recipient.
    /// </summary>
    public enum RecipientType
    {
        /// <summary>
        /// The recipient is a primary (To) recipient.
        /// </summary>
        To,

        /// <summary>
        /// The recipient is a carbon copy (Cc) recipient.
        /// </summary>
        Cc,
        
        /// <summary>
        /// The recipient is a blind carbon copy (Bcc) recipient.
        /// </summary>
        Bcc,

        /// <summary>
        /// The recipient did not successfully receive the message on the previous attempt. 
        /// </summary>
        P1,

        /// <summary>
        /// None.
        /// </summary>
        None
    }
}
