using System;
using System.Runtime.Serialization;

namespace OpenXmlImports.Core
{
    [Serializable]
    public class MaxLengthViolationException : Exception
    {
        public MaxLengthViolationException() { }
        public MaxLengthViolationException(string message) : base(message) { }
        public MaxLengthViolationException(string message, Exception inner) : base(message, inner) { }
        public MaxLengthViolationException(string format, params object[] args)
            : base(string.Format(format, args))
        { }
        protected MaxLengthViolationException(
          SerializationInfo info,
          StreamingContext context)
            : base(info, context) { }
    }
}
