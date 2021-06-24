﻿using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Types
{
    public class ByteType : IType
    {
        public string FriendlyName => "whole number";

        public CellValues DataType => CellValues.Number;

        public object NullSafeGet(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            return byte.Parse(text);
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            if (value != null)
                cellValue.Text = value.ToString();
        }
    }
}
