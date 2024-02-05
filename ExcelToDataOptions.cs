using System.ComponentModel;
using System.Globalization;

#pragma warning disable CS8600, CS8604, CS8601, CS8603, CS8602, CS8618, CS8625

namespace Seamlex.Utilities
{
    ///<summary>Options for controlling Seamlex.Utilities.ExcelToData.</summary>
    [Description("Options for controlling Seamlex.Utilities.ExcelToData.")]
    public class ExcelToDataOptions
    {

        ///<summary>Convert non-CLR types to JSON.</summary>
        [Description("Convert non-CLR types into JSON.")]
        public bool ComplexToJson {get;set;}

        ///<summary>Maximum number of rows to be processed.</summary>
        [Description("Maximum number of rows to be processed.")]
        public int MaxRows {get;set;}

        ///<summary>Whether the first row should be treated as headings.</summary>
        [Description("Whether the first row should be treated as headings.")]
        public bool UseHeadings {get;set;}

        ///<summary>Default table name.</summary>
        [Description("Default table name.")]
        public string DefaultTableName {get;set;}

        ///<summary>Default table column name.</summary>
        [Description("Default table column name.")]
        public string DefaultColumnName {get;set;}

// All dates are UTC format "yyyy-MM-ddTHH:mm:ssZ". Example: 2010-12-04T11:58:00Z
        ///<summary>Date format to read.</summary>
        [Description("Date format to read.")]
        public string SourceDateFormat {get;set;}

        ///<summary>Date/time format to read.</summary>
        [Description("Date/time format to read.")]
        public string SourceDateTimeFormat {get;set;}

        ///<summary>Date format to write.</summary>
        [Description("Number format to write.")]
        public string OutputNumberFormat {get;set;}

        ///<summary>Date format to write.</summary>
        [Description("Date format to write.")]
        public string OutputDateFormat {get;set;}

        ///<summary>Date/time format to write.</summary>
        [Description("Date/time format to write.")]
        public string OutputDateTimeFormat {get;set;}

        ///<summary>Use the OutputDateFormat ahead of the OutputDateTimeFormat.</summary>
        [Description("Use the OutputDateFormat ahead of the OutputDateTimeFormat.")]
        public bool OutputDateOnly {get;set;}

        ///<summary>Check for Excel serial numbers when converting date/time values.</summary>
        [Description("Check for Excel serial numbers when converting date/time values.")]
        public bool DateTimeCheckExcelSerial {get;set;}

        ///<summary>Check for SourceDate[Time]Format when converting date/time values.</summary>
        [Description("Check for SourceDate[Time]Format when converting date/time values.")]
        public bool DateTimeCheckSourceFormat {get;set;}

        ///<summary>Use intelligent date-time fixing when converting date/time values.</summary>
        [Description("Check using intelligent fixing when converting date/time values.")]
        public bool DateTimeCheckFixed {get;set;}

        ///<summary>When using intelligent date-time fixing, check each of these with each culture.</summary>
        [Description("When using intelligent fixing, check each of these with each culture.")]
        public string[] DateTimeCheckFormats {get;set;}

        ///<summary>When using intelligent date-time fixing, will use these cultures.</summary>
        [Description("When using intelligent fixing, will use these cultures.")]
        public List<CultureInfo> DateTimeCheckCultures {get;set;}

        ///<summary>When outputting to Excel, set these columns to OADate numbers not text</summary>
        [Description("When outputting to Excel, set these columns to OADate numbers not text.")]
        public List<string> ColumnsToDateTime{get;set;}

        ///<summary>When outputting to Excel, set these columns to numbers not text</summary>
        [Description("When outputting to Excel, set these columns to numbers not text.")]
        public List<string> ColumnsToNumber{get;set;}

        // CultureInfo[] cultures = { CultureInfo.GetCultureInfo("en-US"), CultureInfo.GetCultureInfo("en-AU") };

        ///<summary>CSV output format.</summary>
        [Description("CSV output format type.")]
        public string CsvFormat {get;set;}

        ///<summary>Wrap all CSV output in \"\" characters.</summary>
        [Description("Wrap all CSV output in \"\" characters.")]
        public bool CsvWrapAll {get;set;}


        ///<summary>CSV newline characters.</summary>
        [Description("CSV newline characters.")]
        public string CsvNewLine {get;set;}


        ///<summary>How to handle text-to-number conversions creating an invalid number.</summary>
        [Description("How to handle text-to-number conversion.")]
        public InvalidNumber InvalidNumberAction {get;set;}
        public enum InvalidNumber
        {
            None,
            Blank,
            Zero,
            FixThenBlank,
            FixThenZero
        }

        ///<summary>How to handle text-to-number conversions exceeding the target size.</summary>
        [Description("How to handle text-to-number conversions exceeding the target size.")]
        public LargeNumber LargeNumberAction {get;set;}
        public enum LargeNumber
        {
            None,
            Blank,
            Zero,
            Max
        }

        ///<summary>How to handle null values in source data.</summary>
        [Description("How to handle null values in source data.")]
        public NullValue NullValueAction {get;set;}
        public enum NullValue
        {
            None,
            Blank,
            Zero,
            Default
        }

        ///<summary>Default values for each simple datatype.</summary>
        [Description("Default values for each simple datatype.")]
        public ExcelToDataBlankValues BlankValues {get;set;}

        ///<summary>Default action for DataTime values.</summary>
        [Description("Default action for DataTime values.")]
        public DateTimeDefault DateTimeDefaultAction {get;set;}
        public enum DateTimeDefault
        {
            Language,
            Lowest,
            Set1900,
            Set1970,
            Set1980,
            Set1990,
            Set2000,
            SetNow
        }

        ///<summary>Excel can decimalise integers imperfectly and therefore when loading we specify how to handle this.</summary>
        [Description("Excel can decimalise integers imperfectly and therefore when loading we specify how to handle this.")]
        public IntegerRounding IntegerRoundingAction {get;set;}
        public enum IntegerRounding
        {
            Ceiling,
            Floor,
            Round
        }
        internal ExcelToDataOptions Clone()
        {
            // 2022-07-14 1.0.1 #2 SNJW added the Clone() method to the options class
            // and the GetDetaultsClone() in the main class
            return (ExcelToDataOptions)this.MemberwiseClone();
        }
    }    
}