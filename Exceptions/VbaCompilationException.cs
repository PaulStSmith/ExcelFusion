using System;
using System.Runtime.Serialization;

namespace ExcelFusion.Exceptions
{
    /// <summary>
    /// Represents an error when compiling a VBA project.
    /// </summary>
    [Serializable]
    internal class VbaCompilationException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="VbaCompilationException"/> class.
        /// </summary>
        public VbaCompilationException() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="VbaCompilationException"/> class.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public VbaCompilationException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="VbaCompilationException"/> class.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        /// <param name="innerException">The exception that is the cause of the current exception, or a null reference (Nothing in Visual Basic) if no inner exception is specified.</param>
        public VbaCompilationException(string message, Exception innerException) : base(message, innerException) { }
    }
}