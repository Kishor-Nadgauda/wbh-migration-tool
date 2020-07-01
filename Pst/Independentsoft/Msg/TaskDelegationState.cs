using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Represents the status types of a delegated task.
    /// </summary>
    public enum TaskDelegationState
    {
        /// <summary>
        /// Specifies that this is not a delegated task or that the task request has been created but not sent. This is also used for a task request message, whether in the owner’s Sent Items folder or the delegate’s Inbox.
        /// </summary>
        NoMatch,

        /// <summary>
        /// Specifies that this is a new task request that has been sent, but the delegate has not yet responded to the task.
        /// </summary>
        OwnNew,

        /// <summary>
        /// Specifies that a task has been accepted. This value should not be in the enumeration.
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
