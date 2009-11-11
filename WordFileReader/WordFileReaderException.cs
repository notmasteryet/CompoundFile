// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.Serialization;

namespace WordFileReader
{
    public class WordFileReaderException : Exception
    {
        public WordFileReaderException(string message)
            : base(message)
        {
        }

        protected WordFileReaderException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

    }
}
