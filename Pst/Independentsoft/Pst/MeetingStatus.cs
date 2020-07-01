using System;

namespace Independentsoft.Pst
{
    /// <summary>
    /// The MeetingStatus enum specifies the status of an appointment or meeting.
    /// </summary>
    public enum MeetingStatus
    {
        NonMeeting,
        Meeting,
        Received,
        CanceledOrganizer,
        Canceled,
        None
    }
}
