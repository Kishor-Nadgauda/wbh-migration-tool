using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Specifies whether the attendee is busy at the time of an appointment on their calendar. The specified status appears in the free/busy view of the calendar.
    /// </summary>
    public enum BusyStatus
    {
        /// <summary>
        /// Free status
        /// </summary>
        Free,

        /// <summary>
        /// Tentative status
        /// </summary>
        Tentative,

        /// <summary>
        /// Busy status
        /// </summary>
        Busy,

        /// <summary>
        /// Out of the office status
        /// </summary>
        OutOfOffice,

        None
    }
}
