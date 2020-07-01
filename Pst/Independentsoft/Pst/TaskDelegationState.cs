using System;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Represents the status types of a delegated task
    /// </summary>
    public enum TaskDelegationState
    {
        /// <summary>
        /// 
        /// </summary>
        NoMatch,

        /// <summary>
        /// 
        /// </summary>
        OwnNew,

        /// <summary>
        /// 
        /// </summary>
        Owned,

        /// <summary>
        /// Specifies that the task has been accepted.
        /// </summary>
        Accepted,

        /// <summary>
        /// Specifies that the task has been declined.
        /// </summary>
        Declined,

        /// <summary>
        /// None.
        /// </summary>
        None
    }
}
