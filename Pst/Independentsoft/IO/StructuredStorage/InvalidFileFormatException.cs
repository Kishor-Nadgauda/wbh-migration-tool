using System;

namespace Independentsoft.IO.StructuredStorage
{
    /// <summary>
    /// The exception that is thrown when an input file or a data stream that is supposed to conform to a certain file format specification is malformed.
    /// </summary>
    public class InvalidFileFormatException : Exception
    {
        /// <summary>
        /// Creates a new instance of the InvalidFileFormatException class. 
        /// </summary>
        public InvalidFileFormatException()
        {
        }

        /// <summary>
        /// Creates a new instance of the InvalidFileFormatException class with the specified error message.  
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public InvalidFileFormatException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates a new instance of the InvalidFileFormatException class with the specified error message and inner exception. 
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        /// <param name="innerException">The exception that is the cause of the current exception. If the innerException parameter is not a null reference (Nothing in Visual Basic), the current exception is raised in a catch block that handles the inner exception.</param>
        public InvalidFileFormatException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}