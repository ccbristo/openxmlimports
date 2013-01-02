using System;
using System.Runtime.Serialization;

namespace OpenXmlImports.Core
{
    [Serializable]
    public class MissingColumnException : Exception
    {
        public MissingColumnException() { }
        public MissingColumnException(string message) : base(message) { }
        public MissingColumnException(string message, params object[] args)
            : base(string.Format(message, args)) { }
        public MissingColumnException(string message, Exception inner) : base(message, inner) { }
        protected MissingColumnException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }
}
