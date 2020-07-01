using System;
using System.Collections.Generic;

namespace Independentsoft.Pst
{
    /// <summary>
    /// Class Appointment.
    /// </summary>
    public class Appointment : Item
    {
        internal Appointment()
        {
        }

        #region Properties

        /// <summary>
        /// Gets the start time.
        /// </summary>
        /// <value>The start time.</value>
        public DateTime StartTime
        {
            get
            {
                return appointmentStartTime;
            }
        }

        /// <summary>
        /// Gets the end time.
        /// </summary>
        /// <value>The end time.</value>
        public DateTime EndTime
        {
            get
            {
                return appointmentEndTime;
            }
        }

        /// <summary>
        /// Gets the location.
        /// </summary>
        /// <value>The location.</value>
        public string Location
        {
            get
            {
                return location;
            }
        }

        /// <summary>
        /// Gets the appointment message class.
        /// </summary>
        /// <value>The appointment message class.</value>
        public string AppointmentMessageClass
        {
            get
            {
                return appointmentMessageClass;
            }
        }

        /// <summary>
        /// Gets the time zone.
        /// </summary>
        /// <value>The time zone.</value>
        public string TimeZone
        {
            get
            {
                return timeZone;
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
        /// Gets the unique identifier.
        /// </summary>
        /// <value>The unique identifier.</value>
        public byte[] Guid
        {
            get
            {
                return guid;
            }
        }

        /// <summary>
        /// Gets the label.
        /// </summary>
        /// <value>The label.</value>
        public long Label
        {
            get
            {
                return label;
            }
        }

        /// <summary>
        /// Gets the duration.
        /// </summary>
        /// <value>The duration.</value>
        public long Duration
        {
            get
            {
                return duration;
            }
        }

        /// <summary>
        /// Gets the busy status.
        /// </summary>
        /// <value>The busy status.</value>
        public BusyStatus BusyStatus
        {
            get
            {
                return busyStatus;
            }
        }

        /// <summary>
        /// Gets the meeting status.
        /// </summary>
        /// <value>The meeting status.</value>
        public MeetingStatus MeetingStatus
        {
            get
            {
                return meetingStatus;
            }
        }

        /// <summary>
        /// Gets the response status.
        /// </summary>
        /// <value>The response status.</value>
        public ResponseStatus ResponseStatus
        {
            get
            {
                return responseStatus;
            }
        }

        /// <summary>
        /// Gets the type of the recurrence.
        /// </summary>
        /// <value>The type of the recurrence.</value>
        public RecurrenceType RecurrenceType
        {
            get
            {
                return recurrenceType;
            }
        }

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
