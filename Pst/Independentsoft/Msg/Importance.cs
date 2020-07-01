using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Indicates the message sender's opinion of the importance of a message.
    /// </summary>
    public enum Importance
    {
        /// <summary>
        /// The message has low importance.
        /// </summary>
        Low,

        /// <summary>
        /// The message has normal importance.
        /// </summary>
        Normal,
        
        /// <summary>
        /// The message has high importance.
        /// </summary>
        High,
        
        /// <summary>
        /// None.
        /// </summary>
        None
    }
}
