using System;
using System.Runtime.Serialization;

namespace ExcelImports.Core
{
    [Serializable]
    public class MissingWorksheetException : Exception
    {
        public MissingWorksheetException() { }
        public MissingWorksheetException(string message) : base(message) { }
        public MissingWorksheetException(string message, Exception inner) : base(message, inner) { }

        public MissingWorksheetException(string message, params object[] args)
            : this(string.Format(message, args))
        { }

        protected MissingWorksheetException(
          SerializationInfo info,
          StreamingContext context)
            : base(info, context) { }
    }
}
