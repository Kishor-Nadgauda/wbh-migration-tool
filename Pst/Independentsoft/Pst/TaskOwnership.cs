using System;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Indicates the ownership state of the task.
    /// </summary>
    public enum TaskOwnership
    {
        /// <summary>
        /// Specifies that task has not yet been assigned to a user.
        /// </summary>
        New,

        /// <summary>
        /// Specifies that task has been delegated to another user.
        /// </summary>
        Delegated,

        /// <summary>
        /// Specifies that task is assigned to the current user.
        /// </summary>
        Own,

        /// <summary>
        /// None.
        /// </summary>
        None
    }
}
