using System;
using System.Runtime.Serialization;

namespace OpenXmlImports.Core
{
    [Serializable]
    public class NullableColumnViolationException : Exception
    {
        public NullableColumnViolationException() { }
        public NullableColumnViolationException(string message) : base(message) { }
        public NullableColumnViolationException(string message, Exception inner) : base(message, inner) { }

        public NullableColumnViolationException(string message, params object[] args)
            : base(string.Format(message, args))
        { }

        protected NullableColumnViolationException(
          SerializationInfo info,
          StreamingContext context)
            : base(info, context) { }
    }
}
