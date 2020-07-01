using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Identifies the sensitivity level assigned to a message item. These levels are arbitrarily set and filtered for, by the user.
    /// </summary>
    public enum Sensitivity
    {
        /// <summary>
        /// The message has the Personal sensitivity setting.
        /// </summary>
        Personal,

        /// <summary>
        /// The message has the Private sensitivity setting.
        /// </summary>
        Private,

        /// <summary>
        /// The message has the Confidential sensitivity setting
        /// </summary>
        Confidential,

        /// <summary>
        /// The message has the Normal sensitivity setting. 
        /// </summary>
        None
    }
}
