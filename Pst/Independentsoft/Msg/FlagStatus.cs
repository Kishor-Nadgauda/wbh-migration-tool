using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Contains the Microsoft Office Outlook follow-up flags for the message.
    /// </summary>
    public enum FlagStatus
    {
        /// <summary>
        /// Complete.
        /// </summary>
        Complete,
        
        /// <summary>
        /// Follow-up is required.
        /// </summary>
        Marked,
        
        /// <summary>
        /// No follow-up has been specified.
        /// </summary>
        None
    }
}
