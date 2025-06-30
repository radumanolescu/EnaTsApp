using System;

namespace Com.Ena.Timesheet.PHD
{
    /// <summary>
    /// Exception class for PHD timesheet processing errors.
    /// </summary>
    public class PhdException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PhdException"/> class.
        /// </summary>
        public PhdException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PhdException"/> class with a specified error message.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public PhdException(string message) : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PhdException"/> class with a specified error message and a reference to the inner exception that is the cause of this exception.
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        /// <param name="inner">The exception that is the cause of the current exception, or a null reference if no inner exception is specified.</param>
        public PhdException(string message, Exception inner) : base(message, inner)
        {
        }
    }
}
