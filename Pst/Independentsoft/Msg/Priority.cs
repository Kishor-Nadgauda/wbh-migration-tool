using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Represents the relative priority of a message.
    /// </summary>
    public enum Priority
    {
        /// <summary>
        /// The message is not urgent.
        /// </summary>
        Low,

        /// <summary>
        /// The message has normal priority.
        /// </summary>
        Normal,

        /// <summary>
        /// The message is urgent.
        /// </summary>
        High,

        /// <summary>
        /// None.
        /// </summary>
        None
    }
}
