using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Represents last action on the message.
    /// </summary>
    public enum LastVerbExecuted
    {
        /// <summary>
        /// Reply has been sent to sender. Value is 102.
        /// </summary>
        ReplyToSender = 102,
        
        /// <summary>
        /// Reply has been sent to all. Value is 103.
        /// </summary>
        ReplyToAll = 103,
        
        /// <summary>
        /// The message has been forwarded. Value is 104.
        /// </summary>
        Forward = 104,
        
        /// <summary>
        /// None.
        /// </summary>
        None = 0
    }
}
