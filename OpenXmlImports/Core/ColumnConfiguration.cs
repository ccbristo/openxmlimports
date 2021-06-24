using System;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlImports.Types;

namespace OpenXmlImports.Core
{
    public class ColumnConfiguration
    {
        public string Name { get; set; }
        public MemberInfo Member { get; set; }
        public NumberingFormat CellFormat { get; set; }
        public bool Required { get; set; }
        public int MaxLength { get; set; }
        public IType Type { get; set; }
        public bool Ignore { get; set; }

        public ColumnConfiguration()
        {
            MaxLength = int.MaxValue;
        }

        // TODO [ccb] Implement these features.
        //public IEnumerable ValidValues { get; set; }
        //public bool ListValidValuesOnError { get; set; }
        //public IComparer OptionComparer { get; set; }

        public object GetValue(object source)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            var value = Member.GetPropertyOrFieldValue(source);
            return value;
        }

        internal void SetValue(object target, object value)
        {
            Member.SetPropertyOrFieldValue(target, value);
        }
    }
}
