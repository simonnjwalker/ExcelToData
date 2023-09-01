using System.ComponentModel;

#pragma warning disable CS8600, CS8604, CS8601, CS8603, CS8602, CS8618, CS8625

namespace Seamlex.Utilities
{
    ///<summary>Values for non-null blanks in Excel data.</summary>
    [Description("Values for non-null blanks in Excel data.")]
    public class ExcelToDataBlankValues
    {
        [Description("Default value for boolean values.")]
        public bool BooleanDefaultValue {get;set;}

        [Description("Default value for Int32 values.")]
        public Int32 Int32DefaultValue {get;set;}

        [Description("Default value for char values.")]
        public char CharDefaultValue {get;set;}

        [Description("Default value for double values.")]
        public double DoubleDefaultValue {get;set;}

        [Description("Default value for single values.")]
        public Single SingleDefaultValue {get;set;}

        [Description("Default value for float values.")]
        public float FloatDefaultValue {get;set;}

        [Description("Default value for string values.")]
        public string StringDefaultValue {get;set;}

        [Description("Default value for decimal values.")]
        public decimal DecimalDefaultValue {get;set;}

        [Description("Default value for long values.")]
        public long LongDefaultValue {get;set;}

        [Description("Default value for short values.")]
        public short ShortDefaultValue {get;set;}

        [Description("Default value for byte values.")]
        public byte ByteDefaultValue {get;set;}

        [Description("Default value for DateTime values.")]
        public DateTime DateTimeDefaultValue {get;set;}

        [Description("Default value for TimeSpan values.")]
        public TimeSpan TimeSpanDefaultValue {get;set;}

        [Description("Default value for object values.")]
        public object ObjectDefaultValue {get;set;}

        [Description("Default value for Guid values.")]
        public Guid GuidDefaultValue {get;set;}

        [Description("Default value for sbyte values.")]
        public sbyte SByteDefaultValue {get;set;}

        [Description("Default value for uint values.")]
        public uint UIntDefaultValue {get;set;}

        [Description("Default value for ulong values.")]
        public ulong ULongDefaultValue {get;set;}

        [Description("Default value for ushort values.")]
        public ushort UShortDefaultValue {get;set;}

        [Description("Default value for char? (nullable char) values.")]
        public char? NullableCharDefaultValue {get;set;}

        [Description("Default value for bool? (nullable bool) values.")]
        public bool? NullableBooleanDefaultValue {get;set;}

        [Description("Default value for int? (nullable int) values.")]
        public int? NullableIntDefaultValue {get;set;}

        [Description("Default value for double? (nullable double) values.")]
        public double? NullableDoubleDefaultValue {get;set;}

        [Description("Default value for decimal? (nullable decimal) values.")]
        public decimal? NullableDecimalDefaultValue {get;set;}

        [Description("Default value for DateTime? (nullable DateTime) values.")]
        public DateTime? NullableDateTimeDefaultValue {get;set;}

        [Description("Default value for float? (nullable float) values.")]
        public float? NullableFloatDefaultValue {get;set;}

        [Description("Default value for long? (nullable long) values.")]
        public long? NullableLongDefaultValue {get;set;}

        [Description("Default value for short? (nullable short) values.")]
        public short? NullableShortDefaultValue {get;set;}

        [Description("Default value for byte? (nullable byte) values.")]
        public byte? NullableByteDefaultValue {get;set;}

        [Description("Default value for sbyte? (nullable sbyte) values.")]
        public sbyte? NullableSByteDefaultValue {get;set;}

        [Description("Default value for uint? (nullable uint) values.")]
        public uint? NullableUIntDefaultValue {get;set;}

        [Description("Default value for ulong? (nullable ulong) values.")]
        public ulong? NullableULongDefaultValue {get;set;}

        [Description("Default value for ushort? (nullable ushort) values.")]
        public ushort? NullableUShortDefaultValue {get;set;}

        [Description("Default value for TimeSpan? (nullable TimeSpan) values.")]
        public TimeSpan? NullableTimeSpanDefaultValue {get;set;}

        [Description("Default value for Guid? (nullable Guid) values.")]
        public Guid? NullableGuidDefaultValue {get;set;}

        public ExcelToDataBlankValues()
        {
            BooleanDefaultValue = false;
            Int32DefaultValue = 0;
            CharDefaultValue = ' ';
            DecimalDefaultValue = 0m;
            DoubleDefaultValue = 0d;
            ObjectDefaultValue = null;
            GuidDefaultValue = Guid.Empty;
            SByteDefaultValue = 0;
            UIntDefaultValue = 0u;
            ULongDefaultValue = 0ul;
            UShortDefaultValue = 0;
            NullableCharDefaultValue = null;
            NullableBooleanDefaultValue = null;
            NullableIntDefaultValue = null;
            NullableDoubleDefaultValue = null;
            NullableDecimalDefaultValue = null;
            NullableDateTimeDefaultValue = null;
            NullableFloatDefaultValue = null;
            NullableLongDefaultValue = null;
            NullableShortDefaultValue = null;
            NullableByteDefaultValue = null;
            NullableSByteDefaultValue = null;
            NullableUIntDefaultValue = null;
            NullableULongDefaultValue = null;
            NullableUShortDefaultValue = null;
        }
    }
}