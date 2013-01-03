using System;
using System.Runtime.Serialization;

namespace OpenXmlImports.Core
{
    [Serializable]
    public class RequiredColumnViolationException : Exception
    {
        public RequiredColumnViolationException() { }
        public RequiredColumnViolationException(string message) : base(message) { }
        public RequiredColumnViolationException(string message, Exception inner) : base(message, inner) { }

        public RequiredColumnViolationException(string message, params object[] args)
            : base(string.Format(message, args))
        { }

        protected RequiredColumnViolationException(
          SerializationInfo info,
          StreamingContext context)
            : base(info, context) { }
    }
}
