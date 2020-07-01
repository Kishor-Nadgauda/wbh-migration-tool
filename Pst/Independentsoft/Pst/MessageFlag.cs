using System;

namespace Independentsoft.Pst
{
    public enum MessageFlag
    {
        /// <summary>
        /// The message is an associated message of a folder. The client or provider has read-only access to this flag. The Read flag is ignored for associated messages, which do not retain a read/unread state. 
        /// </summary>
        Associated,

        /// <summary>
        /// The messaging user sending was the messaging user receiving the message. This flag is meant to be set by the transport provider. 
        /// </summary>
        FromMe,

        /// <summary>
        /// The message has at least one attachment. The client has read-only access to this flag. 
        /// </summary>
        HasAttachment,

        /// <summary>
        /// A nonread report needs to be sent for the message. The client or provider has read-only access to this flag. 
        /// </summary>
        NonReadReportPending,

        /// <summary>
        /// The incoming message arrived over the Internet. It originated either outside the organization or from a source the gateway cannot consider trusted. The client should display an appropriate message to the user. Transport providers set this flag; the client has read-only access. 
        /// </summary>
        OriginInternet,

        /// <summary>
        /// The incoming message arrived over an external link other than X.400 or the Internet. It originated either outside the organization or from a source the gateway cannot consider trusted. The client should display an appropriate message to the user. Transport providers set this flag; the client has read-only access. 
        /// </summary>
        OriginMiscExternal,

        /// <summary>
        /// The incoming message arrived over an X.400 link. It originated either outside the organization or from a source the gateway cannot consider trusted. The client should display an appropriate message to the user. Transport providers set this flag; the client has read-only access. 
        /// </summary>
        OriginX400,

        /// <summary>
        /// The message is marked as having been read. This flag is ignored if the Associated flag is set. 
        /// </summary>
        Read,

        /// <summary>
        /// The message includes a request for a resend operation with a non-delivery report.  
        /// </summary>
        Resend,

        /// <summary>
        /// A read report needs to be sent for the message. The client or provider has read-only access to this flag. 
        /// </summary>
        ReadReportPending,

        /// <summary>
        /// The message is marked for sending. Message store providers set this flag; the client has read-only access. 
        /// </summary>
        Submit,

        /// <summary>
        /// The outgoing message has not been modified since the first time that it was saved; the incoming message has not been modified since it was delivered. 
        /// </summary>
        Unmodified,

        /// <summary>
        /// The message is still being composed. It is saved, but has not been sent. Typically, this flag is cleared after the message is sent.
        /// </summary>
        Unsent
    }
}
