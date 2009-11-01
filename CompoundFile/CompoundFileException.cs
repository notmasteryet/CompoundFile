// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.Serialization;

namespace CompoundFile
{
    /// <summary>
    /// CompoundFile library base exception
    /// </summary>
    public class CompoundFileException : Exception
    {
        /// <summary>
        /// Creates exception.
        /// </summary>
        /// <param name="message">Error message.</param>
        public CompoundFileException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Creates exception.
        /// </summary>
        /// <param name="info">Serialization info.</param>
        /// <param name="context">Streaming context.</param>
        protected CompoundFileException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }
}
