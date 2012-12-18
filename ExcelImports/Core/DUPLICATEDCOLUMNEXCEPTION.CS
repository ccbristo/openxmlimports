using System;
using System.Runtime.Serialization;

namespace ExcelImports.Core
{
    [Serializable]
    public class DuplicatedColumnException : Exception
    {
        public DuplicatedColumnException() { }
        public DuplicatedColumnException(string message) : base(message) { }
        public DuplicatedColumnException(string message, params object[] args)
            : base(string.Format(message, args)) { }
        public DuplicatedColumnException(string message, Exception inner) : base(message, inner) { }
        protected DuplicatedColumnException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        { }
    }
}
