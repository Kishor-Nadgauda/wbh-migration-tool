using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Represents attachment's flags.
    /// </summary>
    public enum AttachmentFlags
    {
        /// <summary>
        /// Indicates that this attachment is not available to HTML rendering applications and should be ignored in MIME processing.
        /// </summary>
        InvisibleInHtml,
        
        /// <summary>
        /// Indicates that this attachment is not available to applications rendering in Rich Text Format (RTF) and should be ignored by MAPI.
        /// </summary>
        InvisibleInRtf,
        
        /// <summary>
        /// None
        /// </summary>
        None
    }
}
