using System;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Identifies the status types of a task item. 
    /// </summary>
    public enum TaskStatus
    {
        /// <summary>
        /// Specifies that the task is not started.
        /// </summary>
        NotStarted,

        /// <summary>
        /// Specifies that the task is in progress.
        /// </summary>
        InProgress,

        /// <summary>
        /// Specifies that the task is completed.
        /// </summary>
        Completed,

        /// <summary>
        /// Specifies that the task is waiting on others.
        /// </summary>
        WaitingOnOthers,

        /// <summary>
        /// Specifies that the task is deferred.
        /// </summary>
        Deferred,

        /// <summary>
        /// None.
        /// </summary>
        None
    }
}
