using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class Task.
    /// </summary>
    public class Task : Item
    {
        internal Task()
        {
        }

        #region Properties

        /// <summary>
        /// Gets the reminder sound file.
        /// </summary>
        /// <value>The reminder sound file.</value>
        public string ReminderSoundFile
        {
            get
            {
                return reminderSoundFile;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is private.
        /// </summary>
        /// <value><c>true</c> if this instance is private; otherwise, <c>false</c>.</value>
        public bool IsPrivate
        {
            get
            {
                return isPrivate;
            }
        }

        /// <summary>
        /// Gets a value indicating whether [reminder override default].
        /// </summary>
        /// <value><c>true</c> if [reminder override default]; otherwise, <c>false</c>.</value>
        public bool ReminderOverrideDefault
        {
            get
            {
                return reminderOverrideDefault;
            }
        }

        /// <summary>
        /// Gets a value indicating whether [reminder play sound].
        /// </summary>
        /// <value><c>true</c> if [reminder play sound]; otherwise, <c>false</c>.</value>
        public bool ReminderPlaySound
        {
            get
            {
                return reminderPlaySound;
            }
        }

        /// <summary>
        /// Gets the owner.
        /// </summary>
        /// <value>The owner.</value>
        public string Owner
        {
            get
            {
                return owner;
            }
        }

        /// <summary>
        /// Gets the delegator.
        /// </summary>
        /// <value>The delegator.</value>
        public string Delegator
        {
            get
            {
                return delegator;
            }
        }

        /// <summary>
        /// Gets the percent complete.
        /// </summary>
        /// <value>The percent complete.</value>
        public double PercentComplete
        {
            get
            {
                return percentComplete;
            }
        }

        /// <summary>
        /// Gets the actual work.
        /// </summary>
        /// <value>The actual work.</value>
        public long ActualWork
        {
            get
            {
                return actualWork;
            }
        }

        /// <summary>
        /// Gets the total work.
        /// </summary>
        /// <value>The total work.</value>
        public long TotalWork
        {
            get
            {
                return totalWork;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is team task.
        /// </summary>
        /// <value><c>true</c> if this instance is team task; otherwise, <c>false</c>.</value>
        public bool IsTeamTask
        {
            get
            {
                return isTeamTask;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is complete.
        /// </summary>
        /// <value><c>true</c> if this instance is complete; otherwise, <c>false</c>.</value>
        public bool IsComplete
        {
            get
            {
                return isComplete;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is recurring.
        /// </summary>
        /// <value><c>true</c> if this instance is recurring; otherwise, <c>false</c>.</value>
        public bool IsRecurring
        {
            get
            {
                return isRecurring;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is all day event.
        /// </summary>
        /// <value><c>true</c> if this instance is all day event; otherwise, <c>false</c>.</value>
        public bool IsAllDayEvent
        {
            get
            {
                return isAllDayEvent;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is reminder set.
        /// </summary>
        /// <value><c>true</c> if this instance is reminder set; otherwise, <c>false</c>.</value>
        public bool IsReminderSet
        {
            get
            {
                return isReminderSet;
            }
        }

        /// <summary>
        /// Gets the reminder time.
        /// </summary>
        /// <value>The reminder time.</value>
        public DateTime ReminderTime
        {
            get
            {
                return reminderTime;
            }
        }

        /// <summary>
        /// Gets the reminder minutes before start.
        /// </summary>
        /// <value>The reminder minutes before start.</value>
        public long ReminderMinutesBeforeStart
        {
            get
            {
                return reminderMinutesBeforeStart;
            }
        }

        /// <summary>
        /// Gets the recurrence pattern description.
        /// </summary>
        /// <value>The recurrence pattern description.</value>
        public string RecurrencePatternDescription
        {
            get
            {
                return recurrencePatternDescription;
            }
        }

        /// <summary>
        /// Gets the recurrence pattern.
        /// </summary>
        /// <value>The recurrence pattern.</value>
        public Independentsoft.Msg.RecurrencePattern RecurrencePattern
        {
            get
            {
                return recurrencePattern;
            }
        }

        /// <summary>
        /// Gets the start date.
        /// </summary>
        /// <value>The start date.</value>
        public DateTime StartDate
        {
            get
            {
                return taskStartDate;
            }
        }

        /// <summary>
        /// Gets the due date.
        /// </summary>
        /// <value>The due date.</value>
        public DateTime DueDate
        {
            get
            {
                return taskDueDate;
            }
        }

        /// <summary>
        /// Gets the date completed.
        /// </summary>
        /// <value>The date completed.</value>
        public DateTime DateCompleted
        {
            get
            {
                return dateCompleted;
            }
        }

        /// <summary>
        /// Gets the status.
        /// </summary>
        /// <value>The status.</value>
        public TaskStatus Status
        {
            get
            {
                return taskStatus;
            }
        }

        /// <summary>
        /// Gets the ownership.
        /// </summary>
        /// <value>The ownership.</value>
        public TaskOwnership Ownership
        {
            get
            {
                return taskOwnership;
            }
        }

        /// <summary>
        /// Gets the state of the delegation.
        /// </summary>
        /// <value>The state of the delegation.</value>
        public TaskDelegationState DelegationState
        {
            get
            {
                return taskDelegationState;
            }
        }

        /// <summary>
        /// Gets the common start time.
        /// </summary>
        /// <value>The common start time.</value>
        public DateTime CommonStartTime
        {
            get
            {
                return commonStartTime;
            }
        }

        /// <summary>
        /// Gets the common end time.
        /// </summary>
        /// <value>The common end time.</value>
        public DateTime CommonEndTime
        {
            get
            {
                return commonEndTime;
            }
        }

        /// <summary>
        /// Gets the flag due by.
        /// </summary>
        /// <value>The flag due by.</value>
        public DateTime FlagDueBy
        {
            get
            {
                return flagDueBy;
            }
        }

        /// <summary>
        /// Gets the companies.
        /// </summary>
        /// <value>The companies.</value>
        public IList<string> Companies
        {
            get
            {
                return companies;
            }
        }

        #endregion
    }
}
