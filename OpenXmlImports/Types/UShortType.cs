﻿using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Types
{
    public class UShortType : IType
    {
        public string FriendlyName
        {
            get { return "whole number"; }
        }

        public CellValues DataType
        {
            get { return CellValues.Number; }
        }

        public object NullSafeGet(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            return ushort.Parse(text);
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            if (value != null)
                cellValue.Text = value.ToString();
        }
    }
}
