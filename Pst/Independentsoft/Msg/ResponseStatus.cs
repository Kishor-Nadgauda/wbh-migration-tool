using System;

namespace Independentsoft.Msg
{
    /// <summary>
    /// Indicates the response to a meeting request.
    /// </summary>
    public enum ResponseStatus
    {
        /// <summary>
        /// The appointment is on the Organizer's calendar or the recipient is the Organizer of the meeting. 
        /// </summary>
        Organized,

        /// <summary>
        /// Meeting tentatively accepted. 
        /// </summary>
        Tentative,

        /// <summary>
        /// Meeting accepted. 
        /// </summary>
        Accepted,

        /// <summary>
        /// Meeting declined. 
        /// </summary>
        Declined,

        /// <summary>
        /// Recipient has not responded. 
        /// </summary>
        NotResponded,

        /// <summary>
        /// The appointment is a simple appointment and does not require a response. 
        /// </summary>
        None
    }
}
