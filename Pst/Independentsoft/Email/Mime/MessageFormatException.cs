using System;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// The exception that is thrown when load or parse message with wrong format.
    /// </summary>
    public class MessageFormatException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the MessageFormatException class. 
        /// </summary>
        public MessageFormatException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the MessageFormatException class with the specified error message.
        /// </summary>
        /// <param name="message">Error description.</param>
        public MessageFormatException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the MessageFormatException class with the specified error message and inner exception. 
        /// </summary>
        /// <param name="message">The message that describes the error</param>
        /// <param name="innerException">The exception that is the cause of the current exception. If the innerException parameter is not a null reference (Nothing in Visual Basic), the current exception is raised in a catch block that handles the inner exception.</param>
        public MessageFormatException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
