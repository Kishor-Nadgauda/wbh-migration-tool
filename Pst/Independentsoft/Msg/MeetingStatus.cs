using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// The MeetingStatus enum specifies the status of an appointment or meeting.
    /// </summary>
    public enum MeetingStatus
    {
        /// <summary>
        /// An Appointment item without attendees has been scheduled. This status can be used to set up holidays on a calendar.
        /// </summary>
        NonMeeting,
        
        /// <summary>
        /// The meeting has been scheduled.
        /// </summary>
        Meeting,
        
        /// <summary>
        /// The meeting request has been received.
        /// </summary>
        Received,
        
        /// <summary>
        /// The scheduled meeting has been cancelled but still appears on the user's calendar.
        /// </summary>
        CanceledOrganizer,
        
        /// <summary>
        /// The scheduled meeting has been cancelled.
        /// </summary>
        Canceled,
        
        /// <summary>
        /// None.
        /// </summary>
        None
    }
}
