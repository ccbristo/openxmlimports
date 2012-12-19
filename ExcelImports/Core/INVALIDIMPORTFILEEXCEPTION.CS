using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace ExcelImports.Core
{
    [Serializable]
    public class InvalidImportFileException : Exception
    {
        public InvalidImportFileException() { }
        public InvalidImportFileException(string message) : base(message) { }
        public InvalidImportFileException(string message, Exception inner) : base(message, inner) { }
        protected InvalidImportFileException(
          SerializationInfo info, StreamingContext context)
            : base(info, context) { }

        public IEnumerable<string> Errors { get; private set; }

        public InvalidImportFileException(IList<string> errors)
        {
            this.Errors = errors;
        }
    }
}
