using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;
using System.ComponentModel;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Collections.Immutable;

#pragma warning disable CS8600, CS8604, CS8601, CS8603, CS8602, CS8618, CS8625, CS8632

namespace Seamlex.Utilities
{
    ///<summary>Simple Excel converter to/from a variety of data sources.</summary>
    [Description("Simple Excel converter to/from a variety of data sources.")]
    public class ExcelToData
    {
        public ExcelToData()
        {
            this.DefaultOptions = BaseDefaults();
        }
        public ExcelToData(ExcelToDataOptions handlerOptions)
        {
            this.DefaultOptions = handlerOptions;
        }

        private ExcelToDataOptions DefaultOptions;

        public void SetOptions(ExcelToDataOptions handlerOptions)
        {
            this.DefaultOptions = handlerOptions;
        }
        public ExcelToDataOptions GetOptions()
        {
            return this.DefaultOptions;
        }
        public ExcelToDataOptions GetOptionsClone()
        {
            // 2022-07-14 1.0.1 #2 SNJW added the Clone() method to the options class
            // this seemed better than adding Newtonsoft
            // I have set the Clone() method to internal as this method is preferred
            return GetOptionsClone(this.DefaultOptions);
        }
        public ExcelToDataOptions GetOptionsClone(ExcelToDataOptions handlerOptions)
        {
            // 2022-07-14 1.0.1 #2 SNJW added the Clone() method to the options class
            // this seemed better than adding Newtonsoft
            // I have set the Clone() method to internal as this method is preferred
            return handlerOptions.Clone();
        }
        private ExcelToDataOptions BaseDefaults()
        {
            return new ExcelToDataOptions()
            {
                MaxRows = 65535,
                UseHeadings = true,
                ComplexToJson = true,
                IntegerRoundingAction = ExcelToDataOptions.IntegerRounding.Round,
                InvalidNumberAction = ExcelToDataOptions.InvalidNumber.Zero,
                LargeNumberAction = ExcelToDataOptions.LargeNumber.Max,
                DefaultTableName = "Sheet1",
                DefaultColumnName = "Column1",
                SourceDateFormat = "d/MM/yyyy",
                SourceDateTimeFormat = "d/MM/yyyy HH:mm:ss aa",
                OutputNumberFormat = "0.00",
                OutputDateFormat = "d/MM/yyyy",
                OutputDateTimeFormat = "d/MM/yyyy HH:mm:ss aa",
                OutputDateOnly = true,
                DateTimeCheckExcelSerial = true,
                DateTimeCheckSourceFormat  = true,
                DateTimeCheckFixed = true,
                DateTimeCheckFormats = new string[]{"yyyy-MM-dd", "dd-MM-yyyy", "MM-dd-yyyy", "dd/MM/yyyy", "MM/dd/yyyy", "yyyy/MM/dd",
                                        "dd.MM.yyyy", "MM.dd.yyyy", "yyyy.MM.dd", "ddMMMyyyy", "dd-MMM-yyyy", "MMM-dd-yyyy",
                                        "dd MMM yyyy", "MMM dd yyyy", "ddd, dd MMM yyyy", "ddd, MMM dd yyyy", "dddd, dd MMM yyyy",
                                        "dddd, MMM dd yyyy", "yyyyMMddTHHmmss", "yyyyMMdd HHmmss", "yyyy-MM-ddTHH:mm:ss",
                                        "yyyy-MM-dd HH:mm:ss"},
                DateTimeCheckCultures = new List<System.Globalization.CultureInfo>{ CultureInfo.GetCultureInfo("en-AU")},
                ColumnsToDateTime = new List<string>(),
                ColumnsToNumber = new List<string>(),
                CsvWrapAll = false,
                CsvFormat = "UTF-8 (Comma delimited)",
                CsvNewLine = "\r\n"
            };
        }

        /// <summary>Last error message if one occurred during processing.</summary>
        [Description("Last error message if one occurred during processing.")]
        public string ErrorMessage = "";


        /// <summary>Convert the first column of the first worksheet in an Excel file into a list of strings.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert the first column of the first worksheet in an Excel file into a list of strings.")]
        public List<string> ToListDataString([Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToListDataString(filePath,"",this.DefaultOptions);
        }
        /// <summary>Convert the first column of a worksheet in an Excel file into a list of strings.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert the first column of a worksheet in an Excel file into a list of strings.")]
        public List<string> ToListDataString([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Name of the Excel worksheet. If blank will use the first.")]string sheetName)
        {
            return this.ToListDataString(filePath,sheetName,this.DefaultOptions);
        }
        /// <summary>Convert the first column of the first worksheet in an Excel file into a list of strings.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert the first column of the first worksheet in an Excel file into a list of strings.")]
        public List<string> ToListDataString([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Name of the Excel worksheet. If blank will use the first.")]string sheetName, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            ErrorMessage = "";
            byte[] byteArray = this.ToExcelBinary(filePath, handlerOptions);
            List<string> listData = new List<string>();
            if(ErrorMessage!="")
                return listData;
            DataSet dataSet = this.ToDataSet(byteArray, handlerOptions);
            if(ErrorMessage!="")
                return listData;
            if(dataSet.Tables.Count==0)
            {
                ErrorMessage = $"No sheets were found in '{filePath}'";
                return listData;
            }
            int sheetNumber = this.GetTableIndex(dataSet,sheetName);
            if(sheetNumber < 0)
                sheetNumber = 0;
            listData.AddRange(this.ToListDataString(dataSet.Tables[sheetNumber], handlerOptions));
            return listData;
        }

        /// <summary>Convert the first column in single worksheet inside an Excel file into a list of integers.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert the first column of the first worksheet in an Excel file into a list of integers.")]
        public List<int> ToListDataInt32([Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToListDataInt32(filePath,"",this.DefaultOptions);
        }
        /// <summary>Convert the first column of a worksheet in an Excel file into a list of integers.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        [Description("Convert the first column of a worksheet in an Excel file into a list of integers.")]
        public List<int> ToListDataInt32([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Name of the Excel worksheet. If blank will use the first.")]string sheetName)
        {
            return this.ToListDataInt32(filePath,sheetName,this.DefaultOptions);
        }
        /// <summary>Convert the first column of a worksheet in an Excel file into a list of integers.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert the first column of a worksheet in an Excel file into a list of integers.")]
        public List<int> ToListDataInt32([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Name of the Excel worksheet. If blank will use the first.")]string sheetName, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            ErrorMessage = "";
            byte[] byteArray = this.ToExcelBinary(filePath, handlerOptions);
            List<int> listData = new List<int>();
            if(ErrorMessage!="")
                return listData;
            DataSet dataSet = this.ToDataSet(byteArray, handlerOptions);
            if(ErrorMessage!="")
                return listData;
            if(dataSet.Tables.Count==0)
            {
                ErrorMessage = $"No sheets were found in '{filePath}'";
                return listData;
            }
            int sheetIndex = 0;
            if(sheetName!="")
                sheetIndex = this.GetTableIndex(dataSet,sheetName);
            if(sheetIndex == -1)
            {
                this.ErrorMessage = $"Excel file '{filePath}' does not contain worksheet '{sheetName}'";
                return listData;
            }
            listData.AddRange(this.ToListDataInt32(dataSet.Tables[sheetIndex], handlerOptions));
            return listData;
        }

        /// <summary>Convert the a worksheet in an Excel file into a list of type 'T'.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert a worksheet in an Excel file into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("Full path and name of the Excel XLSX document.")]string filePath) where T : new()
        {
            return this.ToListData<T>(filePath,"",this.DefaultOptions);
        }
        /// <summary>Convert a worksheet in an Excel file into a list of type 'T'.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will infer it otherwise use the first.</param>
        [Description("Convert a worksheet in an Excel file into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Name of the Excel worksheet. If blank will infer it otherwise use the first.")]string sheetName) where T : new()
        {
            return this.ToListData<T>(filePath,sheetName,this.DefaultOptions);
        }
        /// <summary>Convert a worksheet in an Excel file into a list of type 'T'.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will infer it otherwise use the first.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a worksheet in an Excel file into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Name of the Excel worksheet. If blank will infer it otherwise use the first.")]string sheetName, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new()
        {
            ErrorMessage = "";
            byte[] byteArray = this.ToExcelBinary(filePath, handlerOptions);
            List<T> listData = new();
            if(ErrorMessage!="")
                return listData;
            DataSet dataSet = this.ToDataSet(byteArray, handlerOptions);
            if(ErrorMessage!="")
                return listData;
            if(dataSet.Tables.Count==0)
            {
                ErrorMessage = $"No sheets were found in '{filePath}'";
                return listData;
            }
            int sheetIndex = 0;
            // 2024-01-31 SNJW if the name of the type is not passed in, infer it
            // if(sheetName!="")
            //     sheetIndex = this.GetTableIndex(dataSet,sheetName);
            if(sheetName=="")
            {
                sheetIndex = this.GetTableIndex(dataSet,typeof(T).Name);
            }
            else
            {
                sheetIndex = this.GetTableIndex(dataSet,sheetName);
            }

            if(sheetIndex == -1)
            {
                this.ErrorMessage = $"Excel file '{filePath}' does not contain worksheet '{sheetName}'";
                return listData;
            }
            listData.AddRange(this.ToListData<T>(dataSet.Tables[sheetIndex], handlerOptions));
            return listData;
        }



        /// <summary>Convert a worksheet in an Excel file into a list of type 'T'.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetIndex">Zero=based index Number of the Excel worksheet. If blank will use the first.</param>
        [Description("Convert a worksheet in an Excel file into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("Full path and name of the Excel XLSX document.")]string filePath,[Description("Zero-based index of the Excel worksheet.")]int sheetIndex) where T : new()
        {
            return this.ToListData<T>(filePath,sheetIndex,this.DefaultOptions);
        }
        /// <summary>Convert a worksheet in an Excel file into a list of type 'T'.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetIndex">Zero-based index of the Excel worksheet.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a worksheet in an Excel file into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Zero-based index of the Excel worksheet.")]int sheetIndex, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new()
        {
            ErrorMessage = "";
            DataSet dataSet;
            List<T> listData = new();
            if(Path.GetExtension(filePath).ToLower()==".csv" || Path.GetExtension(filePath).ToLower()==".txt")
            {
                string csvText = "";
                try
                {
                    csvText = System.IO.File.ReadAllText(filePath);
                }
                catch
                {
                    ErrorMessage = $"Could not open CSV file {filePath}";
                }
                if(ErrorMessage!="")
                    return listData;
                string tableName = Path.GetFileNameWithoutExtension(filePath);
                dataSet = this.CsvTextToDataSet(csvText, handlerOptions, tableName);
            }
            else
            {
                byte[] byteArray = this.ToExcelBinary(filePath, handlerOptions);
                if(ErrorMessage!="")
                    return listData;
                dataSet = this.ToDataSet(byteArray, handlerOptions);
            }

            if(ErrorMessage!="")
                return listData;
            if(dataSet.Tables.Count==0)
            {
                ErrorMessage = $"No sheets were found in '{filePath}'";
                return listData;
            }
            if(sheetIndex < 0 || sheetIndex >= dataSet.Tables.Count)
            {
                this.ErrorMessage = $"Excel file '{filePath}' does not contain a worksheet at index {sheetIndex}";
                return listData;
            }
            listData.AddRange(this.ToListData<T>(dataSet.Tables[sheetIndex], handlerOptions));
            return listData;
        }


        /// <summary>Convert the first worksheet inside a binary data Excel file into a list of type 'T'.</summary>
        /// <param name="byteArray">Excel binary data to convert.</param>
        /// 
        [Description("Convert the first worksheet inside a binary data Excel file into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("Excel binary data.")]byte[] byteArray) where T : new()
        {
            return ToListData<T>(byteArray,this.DefaultOptions);
        }
        /// <summary>Convert the first worksheet inside a binary data Excel file into a list of type 'T'.</summary>
        /// <param name="byteArray">Excel binary data to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert the first worksheet inside a binary data Excel file into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("Excel binary data.")]byte[] byteArray, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new()
        {
            ErrorMessage = "";
            List<T> listData = new();
            DataSet dataSet = this.ToDataSet(byteArray, handlerOptions);
            if(ErrorMessage!="")
                return listData;
            if(dataSet.Tables.Count==0)
            {
                ErrorMessage = $"No sheets were found in Excel binary data";
                return listData;
            }
            listData.AddRange(this.ToListData<T>(dataSet.Tables[0], handlerOptions));
            return listData;
        }


        /// <summary>Convert a worksheet inside a binary data Excel file into a list of type 'T'.</summary>
        /// <param name="byteArray">Excel binary data to convert.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        [Description("Convert a worksheet inside a binary data Excel file into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("Excel binary data.")]byte[] byteArray, [Description("Name of the Excel worksheet. If blank will use the first.")]string sheetName) where T : new()
        {
            return ToListData<T>(byteArray,this.DefaultOptions);
        }
        /// <summary>Convert a worksheet inside a binary data Excel file into a list of type 'T'.</summary>
        /// <param name="byteArray">Excel binary data to convert.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a worksheet inside a binary data Excel file into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("Excel binary data.")]byte[] byteArray, [Description("Name of the Excel worksheet. If blank will use the first.")]string sheetName, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new()
        {
            ErrorMessage = "";
            List<T> listData = new();
            DataSet dataSet = this.ToDataSet(byteArray, handlerOptions);
            if(ErrorMessage!="")
                return listData;
            if(dataSet.Tables.Count==0)
            {
                ErrorMessage = $"No sheets were found in Excel binary data";
                return listData;
            }
            int sheetIndex = 0; 
            // 2024-01-31 SNJW if the name of the type is not passed in, infer it
            // if(sheetName!="")
            //     sheetIndex = this.GetTableIndex(dataSet,sheetName);
            if(sheetName=="")
            {
                sheetIndex = this.GetTableIndex(dataSet,typeof(T).Name);
            }
            else
            {
                sheetIndex = this.GetTableIndex(dataSet,sheetName);
            }
            if(sheetIndex == -1)
            {
                this.ErrorMessage = $"Excel binary data does not contain worksheet '{sheetName}'";
                return listData;
            }

            listData.AddRange(this.ToListData<T>(dataSet.Tables[sheetIndex], handlerOptions));
            return listData;
        }
        /// <summary>Convert the first DataTable contained in a DataSet into a list of type 'T'.</summary>
        /// <param name="dataSet">DataSet containing the DataTable to convert.</param>
        [Description("Convert the first DataTable contained in a DataSet into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("DataSet containing the DataTable to convert.")]DataSet dataSet) where T : new()
        {
            return this.ToListData<T>(dataSet,"",this.DefaultOptions);
        }
        /// <summary>Convert a DataTable contained in a DataSet into a list of type 'T'.</summary>
        /// <param name="dataSet">DataSet containing the DataTable to convert.</param>
        /// <param name="tableName">DataTable name to convert. If empty it will convert the first.</param>
        [Description("Convert a DataTable contained in a DataSet into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("DataSet containing the DataTable to convert.")]DataSet dataSet, [Description("DataTable name to convert. If empty it will convert the first.")]string tableName) where T : new()
        {
            return this.ToListData<T>(dataSet,tableName,this.DefaultOptions);
        }
        /// <summary>Convert a DataTable contained in a DataSet into a list of type 'T'.</summary>
        /// <param name="dataSet">DataSet containing the DataTable to convert.</param>
        /// <param name="tableName">DataTable name to convert. If empty it will convert the first.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a DataTable contained in a DataSet into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("DataSet containing the DataTable to convert.")]DataSet dataSet, [Description("DataTable name to convert. If empty it will convert the first.")]string tableName, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new()
        {
            ErrorMessage = "";
            List<T> listData = new();
            if(dataSet.Tables.Count==0)
            {
                ErrorMessage = $"No sheets were found in DataSet";
                return listData;
            }
            int sheetIndex = this.GetTableIndex(dataSet,tableName);
            if(sheetIndex == -1)
            {
                this.ErrorMessage = $"DataSet does not contain table '{tableName}'";
                return listData;
            }
            listData.AddRange(this.ToListData<T>(dataSet.Tables[sheetIndex], handlerOptions));
            return listData;
        }


        /// <summary>Convert a list of type 'T' into columns in a single worksheet inside an Excel file.</summary>
        /// <param name="listData">List of type 'T' to convert.</param>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert a list of type 'T' into columns in a single worksheet inside an Excel file.")]
        public bool ToExcelFile<T>([Description("List of type 'T' convert.")]List<T> listData, [Description("Full path and name of the Excel XLSX document.")]string filePath) where T : new() 
        {
            return this.ToExcelFile<T>(listData,filePath,this.DefaultOptions);
        }
        /// <summary>Convert a list of type 'T' into columns in a single worksheet inside an Excel file.</summary>
        /// <param name="listData">List of type 'T' to convert.</param>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a list of type 'T' into columns in a single worksheet inside an Excel file.")]
        public bool ToExcelFile<T>([Description("List of type 'T' convert.")]List<T> listData, [Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new() 
        {
            ErrorMessage = "";
            var dataTable = this.ToDataTable<T>(listData, handlerOptions);
            byte[] byteArray = this.ToExcelBinary(dataTable, handlerOptions);
            if(ErrorMessage!="")
                return false;
            return this.ToExcelFile(byteArray, filePath, handlerOptions);
        }

        /// <summary>Convert a list of strings into one column in a single worksheet inside an Excel file.</summary>
        /// <param name="listData">List of strings to convert.</param>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert a list of strings into one column in a single worksheet inside an Excel file.")]
        public bool ToExcelFile([Description("List of strings to convert.")]List<string> listData, [Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToExcelFile(listData, filePath, this.DefaultOptions);
        }
        /// <summary>Convert a list of strings into one column in a single worksheet inside an Excel file.</summary>
        /// <param name="listData">List of strings to convert.</param>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a list of strings into one column in a single worksheet inside an Excel file.")]
        public bool ToExcelFile([Description("List of strings to convert.")]List<string> listData, [Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            ErrorMessage = "";
            DataTable dataTable = this.ToDataTable(listData, handlerOptions);
            if(ErrorMessage!="")
                return false;
            byte[] byteArray = this.ToExcelBinary(dataTable, handlerOptions);
            if(ErrorMessage!="")
                return false;
            return this.ToExcelFile(byteArray, filePath, handlerOptions);
        }

        /// <summary>Convert a list of integers into one column in a single worksheet inside an Excel file.</summary>
        /// <param name="listData">List of integers to convert.</param>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert a list of integers into one column in a single worksheet inside an Excel file.")]
        public bool ToExcelFile([Description("List of integers to convert.")]List<int> listData, [Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToExcelFile(listData,"",this.DefaultOptions);
        }
        /// <summary>Convert a list of integers into one column in a single worksheet inside an Excel file.</summary>
        /// <param name="listData">List of integers to convert.</param>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a list of integers into one column in a single worksheet inside an Excel file.")]
        public bool ToExcelFile([Description("List of integers to convert.")]List<int> listData, [Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            ErrorMessage = "";
            DataTable dataTable = this.ToDataTable(listData, handlerOptions);
            if(ErrorMessage!="")
                return false;
            byte[] byteArray = this.ToExcelBinary(dataTable, handlerOptions);
            if(ErrorMessage!="")
                return false;
            return this.ToExcelFile(byteArray, filePath, handlerOptions);
        }

        /// <summary>Convert a DataSet into an Excel file.</summary>
        /// <param name="dataSet">DataSet containing DataTables to convert.</param>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert a DataSet into an Excel file.")]
        public bool ToExcelFile([Description("DataSet containing DataTables to convert.")]DataSet dataSet, [Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToExcelFile(dataSet,filePath,this.DefaultOptions);
        }
        /// <summary>Convert a DataSet into an Excel file.</summary>
        /// <param name="dataSet">DataSet containing DataTables to convert.</param>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a DataSet into an Excel file.")]
        public bool ToExcelFile([Description("DataSet containing DataTables to convert.")]DataSet dataSet, [Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            ErrorMessage = "";
            byte[] byteArray = this.ToExcelBinary(dataSet,handlerOptions);
            if(ErrorMessage!="")
                return false;
            return this.ToExcelFile(byteArray, filePath,handlerOptions);
        }


        /// <summary>Convert a DataTable into an Excel file.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert a DataTable into an Excel file.")]
        public bool ToExcelFile([Description("DataTable to convert.")]DataTable dataTable, [Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToExcelFile(dataTable,filePath,this.DefaultOptions);
        }
        /// <summary>Convert a DataTable into an Excel file.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a DataSet into an Excel file.")]
        public bool ToExcelFile([Description("DataTable to convert.")]DataTable dataTable, [Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable.Copy());
            return this.ToExcelFile(dataSet,filePath,handlerOptions);
        }

        /// <summary>Save an in-memory Excel binary file to a local file.</summary>
        /// <param name="byteArray">Excel binary data to save as a file.</param>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Save an in-memory Excel binary file to a local file.")]
        public bool ToExcelFile([Description("Excel binary data to save as a file.")]byte[] byteArray, [Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToExcelFile(byteArray,filePath,this.DefaultOptions);
        }
        /// <summary>Save an in-memory Excel binary file to a local file.</summary>
        /// <param name="byteArray">Excel binary data to save as a file.</param>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Save an in-memory Excel binary file to a local file.")]
        public bool ToExcelFile([Description("Excel binary data to save as a file.")]byte[] byteArray, [Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            ErrorMessage = "";
            if(ErrorMessage!="")
                return false;
            try
            {
                File.WriteAllBytes(filePath, byteArray);
            }
            catch (IOException ex)
            {
                ErrorMessage = $"Error creating Excel file: '{ex.InnerException}'";
            }
            return ErrorMessage=="";
        }

        /// <summary>Load a local Excel file to in-memory binary data.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Load a local Excel file to in-memory binary data.")]
        public byte[] ToExcelBinary([Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToExcelBinary(filePath,this.DefaultOptions);
        }
        /// <summary>Load a local Excel file to in-memory binary data.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Load a local Excel file to in-memory binary data.")]
        public byte[] ToExcelBinary([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            byte[] result = null;
            ErrorMessage = "";
            try
            {
                result = File.ReadAllBytes(filePath);
            }
            catch (IOException ex)
            {
                ErrorMessage = $"Error loading Excel file: '{ex.InnerException}'";
            }
            return result;
        }

        /// <summary>Convert a list of type 'T' into an in-memory binary Excel file.</summary>
        /// <param name="listData">List of type 'T' to copy into an in-memory Excel file.</param>
        [Description("Convert a list of type 'T' into an in-memory binary Excel file.")]
        public byte[] ToExcelBinary<T>([Description("List of type 'T' to copy into an in-memory Excel file.")]List<T> listData) where T : new()
        {
            return this.ToExcelBinary<T>(listData,this.DefaultOptions);
        }
        /// <summary>Convert a list of type 'T' into an in-memory binary Excel file.</summary>
        /// <param name="listData">List of type 'T' to copy into an in-memory Excel file.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a list of type 'T' into an in-memory binary Excel file.")]
        public byte[] ToExcelBinary<T>([Description("List of type 'T' to copy into an in-memory Excel file.")]List<T> listData, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new()
        {
            var dataTable = this.ToDataTable<T>(listData,handlerOptions);
            return this.ToExcelBinary(dataTable,handlerOptions);
        }

        /// <summary>Convert a DataTable into an in-memory binary Excel file.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        [Description("Convert a DataTable into an in-memory binary Excel file.")]
        public byte[] ToExcelBinary([Description("DataTable to convert.")]DataTable dataTable)
        {
            return this.ToExcelBinary(dataTable,this.DefaultOptions);
        }
        /// <summary>Convert a DataTable into an in-memory binary Excel file.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>        
        [Description("Convert a DataTable into an in-memory binary Excel file.")]
        public byte[] ToExcelBinary([Description("DataTable to convert.")]DataTable dataTable, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);
            return this.ToExcelBinary(dataSet,handlerOptions);
        }

        /// <summary>Convert a DataSet into an in-memory binary Excel file.</summary>
        /// <param name="dataSet">DataSet containing DataTables to convert.</param>
        [Description("Convert a DataSet into an in-memory binary Excel file.")]
        public byte[] ToExcelBinary([Description("DataSet containing DataTables to convert.")]DataSet dataSet)
        {
            return this.ToExcelBinary(dataSet,this.DefaultOptions);
        }
        /// <summary>Convert a DataSet into an in-memory binary Excel file.</summary>
        /// <param name="dataSet">DataSet containing DataTables to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>        
        [Description("Convert a DataSet into an in-memory binary Excel file.")]
        public byte[] ToExcelBinary([Description("DataSet containing DataTables to convert.")]DataSet dataSet, ExcelToDataOptions handlerOptions)
        {
            bool useHeadings = handlerOptions?.UseHeadings ?? this.DefaultOptions.UseHeadings;
            int maxRows = handlerOptions?.MaxRows ?? this.DefaultOptions.MaxRows;
            List<string> columnsToDateTime = handlerOptions?.ColumnsToDateTime ?? new List<string>();
            List<string> columnsToNumber = handlerOptions?.ColumnsToNumber ?? new List<string>();

            string outputDateFormat = handlerOptions?.OutputDateFormat ?? "yyyy-MM-dd";
            string outputDateTimeFormat = handlerOptions?.OutputDateTimeFormat ?? "yyyy-MM-dd hh:mm:ss";
            bool outputDateOnly = handlerOptions?.OutputDateOnly ?? false;

            bool addDateCellStyle = false;
            bool addNumberCellStyle = false;
            if(columnsToDateTime.Count > 0 || columnsToNumber.Count > 0)
            {
                foreach(System.Data.DataTable table in dataSet.Tables)
                {
                    foreach(System.Data.DataColumn column in table.Columns)
                    {
                        if(columnsToDateTime.Contains(table.TableName+'.'+column.ColumnName) || columnsToDateTime.Contains(column.ColumnName))
                        {
                            addDateCellStyle = true;
                        }
                        if(columnsToNumber.Contains(table.TableName+'.'+column.ColumnName) || columnsToNumber.Contains(column.ColumnName))
                        {
                            addNumberCellStyle = true;
                        }

                        if(addDateCellStyle && addNumberCellStyle)
                            break;
                    }
                }
            }

            UInt32Value dateCellFormatIndex = 3;
            UInt32Value numberCellFormatIndex = 7;

            byte[] result = null;
            ErrorMessage = "";
            using (MemoryStream memoryStream = new MemoryStream())
            {
                // try
                // {
                    // Create a new spreadsheet document
                    using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                    {
                        // Add a WorkbookPart to the document
                        WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();

                        // Add a WorksheetPart to the WorkbookPart
                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());

                        // Add Sheets to the Workbook
                        Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                        int tableid = 0;

                        // we have to do some work to format dates
                        if(addDateCellStyle || addNumberCellStyle)
                        {
                            this.AddBaseStylesheet(spreadsheet,handlerOptions);
                        }

                        foreach(System.Data.DataTable dataTable in dataSet.Tables)
                        {
                            // Append a new worksheet and associate it with the workbook
                            tableid++;
                            string tableName = (dataTable.TableName ?? "");
                            if(!this.IsValidWorksheetName(tableName))
                            {
                                tableName = "Sheet"+tableid.ToString();
                                if(dataSet.Tables.Cast<DataTable>().Any(table => table.TableName == tableName))
                                    tableName = System.Guid.NewGuid().ToString().Replace("-","").Substring(0,31);
                            }
                            Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = new DocumentFormat.OpenXml.UInt32Value((uint)tableid), Name = tableName };
                            sheets.Append(sheet);

                            // Get the sheetData cell table
                            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                            // Add the header row
                            if(useHeadings)
                            {
                                Row headerRow = new Row();
                                foreach (DataColumn column in dataTable.Columns)
                                {
                                    Cell cell = new Cell();
                                    cell.DataType = CellValues.String;
                                    cell.CellValue = new CellValue(column.ColumnName);
                                    headerRow.AppendChild(cell);
                                }
                                /*

                                if (cellValue.StartsWith("="))
                                {
                                    cell.CellFormula = new CellFormula(cellValue);
                                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                                }
                                else
                                {
                                    cell.DataType = CellValues.String;
                                    cell.CellValue = new CellValue(cellValue);
                                }

                                */

                                sheetData.AppendChild(headerRow);
                            }

                            // Add the data rows
                            int rowcount = 0;
                            foreach (DataRow dataRow in dataTable.Rows)
                            {
                                if(rowcount>=(maxRows-1))
                                    break;
                                Row newRow = new Row();
                                foreach (DataColumn column in dataTable.Columns)
                                {
                                    Cell cell = new Cell();
                                    cell.DataType = CellValues.String;


                                    if(columnsToDateTime.Contains(tableName+'.'+column.ColumnName) || columnsToDateTime.Contains(column.ColumnName))
                                    {
                                        DateTime? checkdate = this.ParseExcelDateTimeNullable(dataRow[column].ToString(),handlerOptions);

                                        if(checkdate != null)
                                        {
                                            cell.DataType = CellValues.Number;
                                            cell.CellValue = new CellValue(((DateTime)checkdate).ToOADate());
                                            cell.StyleIndex = dateCellFormatIndex;
                                        }
                                        else
                                        {
                                            cell.CellValue = new CellValue(dataRow[column].ToString());
                                        }
                                    }
                                    else if(columnsToNumber.Contains(tableName+'.'+column.ColumnName) || columnsToNumber.Contains(column.ColumnName))
                                    {
                                        cell.DataType = CellValues.Number;
                                        cell.CellValue = new CellValue(dataRow[column].ToString());
                                        cell.StyleIndex = numberCellFormatIndex;
                                    }
                                    else
                                    {
                                        cell.CellValue = new CellValue(dataRow[column].ToString());
                                    }
                                    newRow.AppendChild(cell);
                                }
                                sheetData.AppendChild(newRow);
                                rowcount++;
                            }

                        }

                        // Save the changes
                        // workbookPart.Workbook.Save();
                        spreadsheet.Close();
                        result = memoryStream.ToArray();
                    }
                // }
                // catch (Exception ex)
                // {
                //     ErrorMessage = $"Error creating Excel file: '{ex.Message}'";
                // }
            }
            return result;
        }

        public void AddBaseStylesheet(SpreadsheetDocument spreadsheet)
        {
            this.AddBaseStylesheet(spreadsheet,this.DefaultOptions);
        }

        public void AddBaseStylesheet(SpreadsheetDocument spreadsheet, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            WorkbookStylesPart stylesPart = spreadsheet?.WorkbookPart?.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = this.CreateBaseStylesheet(handlerOptions);
            spreadsheet.WorkbookPart.WorkbookStylesPart.Stylesheet.Save();
        }
        

        // derived from: https://jason-ge.medium.com/create-excel-using-openxml-in-net-6-3b601ddf48f7
        private Stylesheet CreateBaseStylesheet()
        {
            return this.CreateBaseStylesheet(this.DefaultOptions);
        }
        private Stylesheet CreateBaseStylesheet([Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            Stylesheet stylesheet = new Stylesheet();

            uint baseDateTimeFormatId = 164;
            uint baseNumberFormatId = 165;
            bool outputDateOnly = handlerOptions?.OutputDateOnly ?? false;
            string outputDateFormat = handlerOptions?.OutputDateFormat ?? "yyyy-mm-dd";
            if(!outputDateOnly)
                outputDateFormat = handlerOptions?.OutputDateTimeFormat ?? "yyyy-mm-dd hh:mm:ss";
            string outputNumberFormat = handlerOptions?.OutputNumberFormat ?? "#.##";


            var numberingFormats = new NumberingFormats();
            numberingFormats.Append(new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(baseDateTimeFormatId),
                FormatCode = StringValue.FromString(outputDateFormat)
            });
            numberingFormats.Append(new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(baseNumberFormatId),
                FormatCode = StringValue.FromString(outputNumberFormat)
            });
            numberingFormats.Count = UInt32Value.FromUInt32((uint)numberingFormats.ChildElements.Count);

            var fonts = new Fonts();
            fonts.Append(new DocumentFormat.OpenXml.Spreadsheet.Font()  // Font index 0 - default
            {
                FontName = new FontName { Val = StringValue.FromString("Calibri") },
                FontSize = new FontSize { Val = DoubleValue.FromDouble(11) }
            });
            fonts.Append(new DocumentFormat.OpenXml.Spreadsheet.Font()  // Font index 1
            {
                FontName = new FontName { Val = StringValue.FromString("Arial") },
                FontSize = new FontSize { Val = DoubleValue.FromDouble(11) },
                Bold = new Bold()
            });
            fonts.Count = UInt32Value.FromUInt32((uint)fonts.ChildElements.Count);
            var fills = new Fills();
            fills.Append(new Fill()
            {
                PatternFill = new PatternFill { PatternType = PatternValues.None }
            });
            fills.Append(new Fill()
            {
                PatternFill = new PatternFill { PatternType = PatternValues.Gray125 }
            });
            fills.Append(new Fill()
            {
                PatternFill = new PatternFill { 
                    PatternType = PatternValues.Solid, 
                    ForegroundColor = TranslateForeground(System.Drawing.Color.LightBlue),
                    BackgroundColor = new BackgroundColor { Rgb = TranslateForeground(System.Drawing.Color.LightBlue).Rgb }
                }
            });
            fills.Append(new Fill()
            {
                PatternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = TranslateForeground(System.Drawing.Color.LightSkyBlue),
                    BackgroundColor = new BackgroundColor { Rgb = TranslateForeground(System.Drawing.Color.LightBlue).Rgb }
                }
            });
            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);
            var borders = new Borders();
            borders.Append(new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Append(new Border
            {
                LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin },
                RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Append(new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);
            var cellStyleFormats = new CellStyleFormats();
            cellStyleFormats.Append(new CellFormat  // Cell style format index 0: no format
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0
            });
            cellStyleFormats.Count = UInt32Value.FromUInt32((uint)cellStyleFormats.ChildElements.Count);
            var cellFormats = new CellFormats();
            cellFormats.Append(new CellFormat());    // Cell format index 0
            cellFormats.Append(new CellFormat   // CellFormat index 1
            {
                NumberFormatId = 14,        // 14 = 'mm-dd-yy'. Standard Date format;
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // Cell format index 2: Standard Number format with 2 decimal placing
            {
                NumberFormatId = 4,        // 4 = '#,##0.00';
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // Cell format index 3
            {
                NumberFormatId = baseDateTimeFormatId,        // 164 = custom (see above)
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // Cell format index 4
            {
                NumberFormatId = 3, // 3   #,##0
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 5
            {
                NumberFormatId = 4, // 4   #,##0.00
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // Cell format index 6
            {
                NumberFormatId = 10,    // 10  0.00 %,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // Cell format index 7
            {
                NumberFormatId = baseNumberFormatId,    // Format cellas 4 digits. If less than 4 digits, prepend 0 in front
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // Cell format index 8: Cell header
            {
                NumberFormatId = 49,
                FontId = 1,
                FillId = 3,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center }
            });
            cellFormats.Count = UInt32Value.FromUInt32((uint)cellFormats.ChildElements.Count);
            stylesheet.Append(numberingFormats);
            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);

            var css = new CellStyles();
            css.Append(new CellStyle
            {
                Name = StringValue.FromString("Normal"),
                FormatId = 0,
                BuiltinId = 0
            });
            css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
            stylesheet.Append(css);

            var dfs = new DifferentialFormats { Count = 0 };
            stylesheet.Append(dfs);
            var tss = new TableStyles
            {
                Count = 0,
                DefaultTableStyle = StringValue.FromString("TableStyleMedium9"),
                DefaultPivotStyle = StringValue.FromString("PivotStyleLight16")
            };
            stylesheet.Append(tss);

            return stylesheet;
        }
        private Columns AutoSizeCells(SheetData sheetData)
        {
            var maxColWidth = GetMaxCharacterWidth(sheetData);

            Columns columns = new Columns();
            //this is the width of my font - yours may be different
            double maxWidth = 7;
            foreach (var item in maxColWidth)
            {
                //width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
                double width = Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256;
                Column col = new Column() { BestFit = true, Min = (UInt32)(item.Key + 1), Max = (UInt32)(item.Key + 1), CustomWidth = true, Width = (DoubleValue)width };
                columns.Append(col);
            }

            return columns;
        }

        private ForegroundColor TranslateForeground(System.Drawing.Color fillColor)
        {
            return new ForegroundColor()
            {
                Rgb = new HexBinaryValue()
                {
                    Value =
                            System.Drawing.ColorTranslator.ToHtml(
                            System.Drawing.Color.FromArgb(
                                fillColor.A,
                                fillColor.R,
                                fillColor.G,
                                fillColor.B)).Replace("#", "")
                }
            };
        }

        private Dictionary<int, int> GetMaxCharacterWidth(SheetData sheetData)
        {
            //iterate over all cells getting a max char value for each column
            Dictionary<int, int> maxColWidth = new Dictionary<int, int>();
            var rows = sheetData.Elements<Row>();
            UInt32[] numberStyles = new UInt32[] { 5, 6, 7, 8 }; //styles that will add extra chars
            UInt32[] boldStyles = new UInt32[] { 1, 2, 3, 4, 6, 7, 8 }; //styles that will bold
            foreach (var r in rows)
            {
                var cells = r.Elements<Cell>().ToArray();

                //using cell index as my column
                for (int i = 0; i < cells.Length; i++)
                {
                    var cell = cells[i];
                    var cellValue = cell.CellValue == null ? cell.InnerText : cell.CellValue.InnerText;
                    var cellTextLength = cellValue.Length;

                    if (cell.StyleIndex != null && numberStyles.Contains(cell.StyleIndex))
                    {
                        int thousandCount = (int)Math.Truncate((double)cellTextLength / 4);

                        //add 3 for '.00' 
                        cellTextLength += (3 + thousandCount);
                    }

                    if (cell.StyleIndex != null && boldStyles.Contains(cell.StyleIndex))
                    {
                        //add an extra char for bold - not 100% acurate but good enough for what i need.
                        cellTextLength += 1;
                    }

                    if (maxColWidth.ContainsKey(i))
                    {
                        var current = maxColWidth[i];
                        if (cellTextLength > current)
                        {
                            maxColWidth[i] = cellTextLength;
                        }
                    }
                    else
                    {
                        maxColWidth.Add(i, cellTextLength);
                    }
                }
            }

            return maxColWidth;
        }        

        /// <summary>Convert an Excel (XLSX/CSV) file into a DataSet.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert an Excel (XLSX/CSV) file into a DataSet.")]
        public DataSet ToDataSet([Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToDataSet(filePath,this.DefaultOptions);
        }
        /// <summary>Convert an Excel (XLSX/CSV) file into a DataSet.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert an Excel (XLSX/CSV) file into a DataSet.")]
        public DataSet ToDataSet([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            ErrorMessage = "";
            DataSet dataSet;
            if(Path.GetExtension(filePath).ToLower()==".csv" || Path.GetExtension(filePath).ToLower()==".txt")
            {
                string csvText = "";
                try
                {
                    csvText = System.IO.File.ReadAllText(filePath);
                }
                catch
                {
                    ErrorMessage = $"Could not open CSV file {filePath}";
                }
                string tableName = Path.GetFileNameWithoutExtension(filePath);
                dataSet = this.CsvTextToDataSet(csvText, handlerOptions, tableName);
            }
            else
            {
                byte[] byteArray = this.ToExcelBinary(filePath, handlerOptions);
                dataSet = this.ToDataSet(byteArray, handlerOptions);
            }
            return dataSet;
        }

        /// <summary>Convert an in-memory binary Excel file into a DataSet.</summary>
        /// <param name="byteArray">In-memory Excel file to convert.</param>
        [Description("Convert an in-memory binary Excel file into a DataSet.")]
        public DataSet ToDataSet([Description("In-memory Excel file to convert.")]byte[] byteArray)
        {
            return this.ToDataSet(byteArray,this.DefaultOptions);
        }
        /// <summary>Convert an in-memory binary Excel file into a DataSet.</summary>
        /// <param name="byteArray">In-memory Excel file to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert an in-memory binary Excel file into a DataSet.")]
        public DataSet ToDataSet([Description("In-memory Excel file to convert.")]byte[] byteArray, ExcelToDataOptions handlerOptions)
        {
            bool useHeadings = handlerOptions?.UseHeadings ?? this.DefaultOptions.UseHeadings;
            int maxRows = handlerOptions?.MaxRows ?? this.DefaultOptions.MaxRows;
            string defaultColumnName = handlerOptions?.DefaultColumnName ?? this.DefaultOptions.DefaultColumnName;
            ErrorMessage = "";
            DataSet output = new DataSet();
            if(byteArray==null)
            {
                ErrorMessage = "No binary data was retrieved";
                return output;
            }
            List<DataTable> tables = new List<DataTable>();
            using (MemoryStream memoryStream = new MemoryStream(byteArray))
            {
                // try
                // {
                    // Open the excel file
                    using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(memoryStream, false))
                    {
                        // Get the workbook part
                        WorkbookPart workbookPart = spreadsheet.WorkbookPart;

                        // Get the sheets
                        Sheets sheets = workbookPart.Workbook.Sheets;

                        // Iterate through the sheets
                        foreach (Sheet sheet in sheets.Cast<Sheet>())
                        {
                            // Get the worksheet part
                            WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));

                            // Create a new DataTable
                            DataTable dataTable = new DataTable(sheet.Name);

                            // Get the sheet data
                            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                            // Add the columns
                            int columnIndex = 0;
                            int columnTotal = 0;
                            int skipFirst = 1;

                            // 2024-01-31 SNJW there is an issue with blank values
                            // so need to get the index of the last value in the first row and fill in the gaps
                            List<string> columnNames = GetColumnNames(workbookPart,sheetData.Descendants<Row>().First(),useHeadings);
                            foreach(var columnName in columnNames)
                                dataTable.Columns.Add(new DataColumn(columnName, typeof(string)));
                            columnTotal = dataTable.Columns.Count;





            //                 if(useHeadings)
            //                 {
            //                     Row headerRow = sheetData.Descendants<Row>().First();
            //                     if(headerRow.HasChildren==true)
            //                     {
            //                         foreach (Cell cell in headerRow.Descendants<Cell>())
            //                         {
            // // 2022-09-01 1.0.2 #3 SNJW if the column is merged or blank then give it a placeholder name
            // // this is okay as it will not align with 
            //                             string columnName = (GetCellValue(workbookPart, cell) ?? "").ToString();
            //                             if(!this.CanAddColumn(dataTable,columnName))
            //                             {
            //                                 if(columnIndex==0)
            //                                 {
            //                                     columnName = defaultColumnName;
            //                                 }
            //                                 else
            //                                 {
            //                                     // columnName = System.Guid.NewGuid().ToString().Replace("-","").ToLower().Substring(0,31);
            //                                     columnName = System.Guid.NewGuid().ToString().Replace("-","").ToLower()[..31];                                                
            //                                 }
            //                             }
            //                             dataTable.Columns.Add(new DataColumn(columnName, typeof(string)));
            //                             columnIndex++;
            //                             columnTotal++;
            //                         }
            //                     }
            //                     else
            //                     {
            //                         dataTable.Columns.Add(new DataColumn(defaultColumnName, typeof(string)));
            //                         columnIndex++;
            //                     }
            //                 }
            //                 else
            //                 {
            //                     skipFirst = 0;
            //                     Row firstRow = sheetData.Descendants<Row>().First();
            //                     for(int i = 0; i<firstRow.Descendants().Count(); i++)
            //                     {
            //                         string columnName = defaultColumnName.TrimEnd('1') + i.ToString().PadLeft(3,'0');
            //                         dataTable.Columns.Add(new DataColumn(columnName, typeof(string)));
            //                         columnIndex++;
            //                     }
            //                 }


                            

                            // Add the data rows
                            int rowcount = 0;
//                            foreach (Row row in sheetData.Elements<Row>().Skip(skipFirst))
// row.Descendants<Cell>().ElementAt(i)
                            foreach (Row row in sheetData.Descendants<Row>().Skip(skipFirst))
                            {
                                rowcount++;
                                if(rowcount>=maxRows)
                                    break;

                                DataRow dataRow = dataTable.NewRow();
                                columnIndex = 0;

                                // 2024-01-31 SNJW there is an issue with blank cells where OpenXML simply will not detect them
                                int descendantCount = row.Descendants<Cell>().Count();
                                if(descendantCount==0 || descendantCount != columnTotal)
                                    for (int i = 0; i < dataRow.ItemArray.Length; i++)
                                        dataRow[i]="";

                                if(descendantCount == columnTotal)
                                {
                                    // this is the original code
                                    for(int i = 0; i < columnTotal; i++) // 
                                    {
                                        Cell cell = null;
                                        try
                                        {
                                            cell = row.Descendants<Cell>().ElementAt(i);
                                        }
                                        catch
                                        {

                                        }
                                        if(cell == null)
                                        {
                                            dataRow[i] = "";
                                        }
                                        else
                                        {
                                            string cellValue = (GetCellValue(workbookPart, cell) ?? "").ToString();
                                            dataRow[i] = cellValue;
                                        }
                                    }                                    
                                }
                                else if(descendantCount>0)
                                {
                                    // go through each of the cells that exist, find the index of these and insert it in the correct spot
                                    for(int i = 0; i < descendantCount; i++) // 
                                    {
                                        Cell cell = null;
                                        try
                                        {
                                            cell = row.Descendants<Cell>().ElementAt(i);
                                        }
                                        catch
                                        {

                                        }
                                        if(cell != null)
                                        {
                                            string cellValue = (GetCellValue(workbookPart, cell) ?? "").ToString();
                                            int cellIndex = GetCellIndex(cell.CellReference);
                                            dataRow[cellIndex] = cellValue;
                                        }
                                    }
                                }








// //                                foreach (Cell cell in row.Elements<Cell>())
//                                 for(int i = 0; i < columnTotal; i++) // 
//                                 {
//                                     Cell cell = null;
//                                     try
//                                     {
//                                         cell = row.Descendants<Cell>().ElementAt(i);
//                                     }
//                                     catch
//                                     {

//                                     }
//                                     if(cell == null)
//                                     {
//                                         dataRow[columnIndex] = "";
//                                     }
//                                     else
//                                     {
//                                         string cellValue = (GetCellValue(workbookPart, cell) ?? "").ToString();
//                                         dataRow[columnIndex] = cellValue;
//                                     }
//                                     columnIndex++;
//                                 }



                                dataTable.Rows.Add(dataRow);
                            }

                            // Add the DataTable to the result
                            tables.Add(dataTable);
                        }
                    }
                // }
                // catch (Exception ex)
                // {
                //     ErrorMessage = $"Error reading Excel file: '{ex.InnerException}'";
                // }
            }
            foreach(var table in tables)
                output.Tables.Add(table);
            return output;
        }

        private List<string> GetColumnNames(WorkbookPart workbookPart, Row row, bool useHeadings)
        {
            // we need to find the index value of each item in the top row of the XLSX
            int descendantCount = row.Descendants<Cell>().Count();
            List<string> output = new();

            if(descendantCount==0)
                return output;
            
            int lastCellIndex = GetCellIndex(row.Descendants<Cell>().Last().CellReference);

            // if these align 1-to-1, we can cycle through them or use default names
            if(descendantCount == ( lastCellIndex + 1))
            {
                if(useHeadings)
                {
                    // use the headings if they are unique
                    for(int i = 0; i < descendantCount; i++) // 
                    {
                        Cell cell = null;
                        try
                        {
                            cell = row.Descendants<Cell>().ElementAt(i);
                        }
                        catch
                        {

                        }
                        string headingName = $"Column{i+1}";
                        if(cell != null)
                        {
                            string cellValue = (GetCellValue((WorkbookPart)workbookPart, cell) ?? "").ToString();
                            if(cellValue=="")
                            {
                                cellValue = headingName;
                            }
                            else if(cellValue.Length==1)
                            {
                                if(!char.IsLetterOrDigit(cellValue[0]))
                                    cellValue = headingName;
                            }
                            else
                            {
                                string fieldCheck = "";
                                for(int j = 0; j < cellValue.Length; j++)
                                {
                                    if(j == 0)
                                    {
                                        if(char.IsLetter(cellValue[0]))
                                        {
                                            fieldCheck = cellValue[0].ToString();
                                        }
                                        else
                                        {
                                            fieldCheck = "_";
                                        }
                                    }
                                    else
                                    {
                                        if(!char.IsLetterOrDigit(cellValue[j]))
                                        {
                                            fieldCheck += "_";
                                        }
                                        else
                                        {
                                            fieldCheck += cellValue[j];
                                        }
                                    }
                                }
                                if(fieldCheck.Length>64)
                                    fieldCheck = fieldCheck.Substring(0,64);
                                cellValue = fieldCheck;
                            }

                            if(output.Exists(x => x.Equals(cellValue)))
                                cellValue = headingName;
                            headingName = cellValue;
                        }

                        if(output.Exists(x => x.Equals(headingName)))
                            headingName = System.Guid.NewGuid().ToString().Replace("-","").ToLower();
                        output.Add(headingName);
                    }
                }
                else
                {
                    // use the default column names
                    for (int i = 0; i < descendantCount; i++)
                        output.Add($"Column{i+1}");
                }
            }
            return output;
        }

        private string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            string cellValue = cell?.CellValue?.InnerText;
            if (cell?.DataType != null && cell?.DataType?.Value == CellValues.SharedString)
            {
                SharedStringItem sharedString = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cellValue));
                cellValue = sharedString.Text.Text;
            }
            return cellValue;
        }

        private int GetCellIndex(string cellReference)
        {
            // Remove any numeric characters to isolate the column part
            string columnPart = string.Empty;
            foreach (char c in cellReference)
            {
                if (char.IsLetter(c))
                {
                    columnPart += c;
                }
            }

            // Convert column letters to zero-based column index
            int columnIndex = 0;
            for (int i = 0; i < columnPart.Length; i++)
            {
                columnIndex *= 26;
                columnIndex += (columnPart[i] - 'A' + 1);
            }

            return columnIndex - 1; // Subtract 1 to make it zero-based
        }


    private string SanitiseFieldName(string input)
    {
        if (string.IsNullOrEmpty(input))
            return System.Guid.NewGuid().ToString().Replace("-","").ToLower();

        if(input.Length==1 && char.IsLetter(input[0]))
        {
            return input;
        }

        var sb = new System.Text.StringBuilder();
        bool isFirstChar = true;
        foreach (char c in input)
        {
            // First character must be a letter or an underscore
            if (isFirstChar)
            {
                if (char.IsLetter(c))
                {
                    sb.Append(c);
                }
                else
                {
                    sb.Append('_');
                }
                isFirstChar = false;
            }
            else
            {
                if (char.IsLetterOrDigit(c) || c == '_')
                {
                    sb.Append(c);
                }
                // Optionally, handle other characters like spaces, hyphens, etc., here
                // For example, convert them to underscores
                else
                {
                    sb.Append('_');
                }
            }
        }

//        return System.Guid.NewGuid().ToString().Replace("-","").ToLower();
        string result = sb.ToString();
        return result;
    }


/*
Important:  Worksheet names cannot:
Be blank.
Contain more than 31 characters.
Contain any of the following characters: / \ ? * : [ ]
For example, 02/17/2016 would not be a valid worksheet name, but 02-17-2016 would work fine.
Begin or end with an apostrophe ('), but they can be used in between text or numbers in a name.
Be named "History". This is a reserved word Excel uses internally.
*/

        /// <summary>Check whether a name can be an Excel worksheet.
        // // Worksheet names cannot:
        // // Be blank.
        // // Contain more than 31 characters.
        // // Contain any of the following characters: / \ ? * : [ ]
        // // For example, 02/17/2016 would not be a valid worksheet name, but 02-17-2016 would work fine.
        // // Begin or end with an apostrophe ('), but they can be used in between text or numbers in a name.
        // // Be named "History". This is a reserved word the MS Excel uses internally.
        /// </summary>
        /// <param name="testName">Text to test whether this can be a valid Excel worksheet name.</param>
        [Description("Check whether a name can be an Excel worksheet.")]
        public bool IsValidWorksheetName([Description("Text to test whether this can be a valid Excel worksheet name.")]string testName)
        {
            if(testName=="")
                return false;
            if(testName.Length>31)
                return false;
            if(new[] { '/', '\\', '?', '*', ':', '[', ']' }.Any(testName.Contains))
                return false;
            if(testName.StartsWith("'") || testName.EndsWith("'"))
                return false;
            if(testName.ToLower() == "history")
                return false;
            return true;
        }


        /// <summary>Perform a case-insensitive check whether a column name exists in a datatable.</summary>
        /// <param name="dataTable">DataTable to check.</param>
        /// <param name="testName">Text to test whether this can added as a valid column in this DataTable.</param>
        [Description("Check whether a name can be an Excel worksheet.")]
        public bool CanAddColumn([Description("DataTable to check.")]DataTable dataTable, [Description("Text to test whether this can added as a valid column in this DataTable.")]string testName)
        {
            testName = testName.ToLower().Trim();
            if(testName=="")
                return false;
            for(int i = 0; i < dataTable.Columns.Count; i++)
            {
                if(dataTable.Columns[i].ColumnName.ToLower().Trim() == testName)
                    return false;
            }
            return true;
        }

        /// <summary>Convert the first worksheet inside an Excel file into a DataTable.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert the first worksheet inside an Excel file into a DataTable.")]
        public DataTable ToDataTable([Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToDataTable(filePath,0,this.DefaultOptions);
        }
        /// <summary>Convert the first worksheet inside an Excel file into a DataTable.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a single worksheet inside an Excel file into a DataTable.")]
        public DataTable ToDataTable([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            var dataSet = this.ToDataSet(filePath,handlerOptions);
            if(dataSet.Tables.Count == 0)
                return new DataTable();
            return dataSet.Tables[0];
        }


        /// <summary>Convert a single worksheet inside an Excel file into a DataTable.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        [Description("Convert a single worksheet inside an Excel file into a DataTable.")]
        public DataTable ToDataTable([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Name of the Excel worksheet. If blank will use the first.")]string sheetName)
        {
            return this.ToDataTable(filePath,sheetName,this.DefaultOptions);
        }

        /// <summary>Convert a single worksheet inside an Excel file into a DataTable.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetIndex">Zero-based index of the Excel worksheet.</param>
        [Description("Convert a single worksheet inside an Excel file into a DataTable.")]
        public DataTable ToDataTable([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Zero-based index of the Excel worksheet.")]int sheetIndex)
        {
            return this.ToDataTable(filePath,sheetIndex,this.DefaultOptions);
        }
        /// <summary>Convert a single worksheet inside an Excel file into a DataTable.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetIndex">Zero-based index of the Excel worksheet.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a single worksheet inside an Excel file into a DataTable.")]
        public DataTable ToDataTable([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Zero-based index of the Excel worksheet.")]int sheetIndex, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            var dataSet = this.ToDataSet(filePath,handlerOptions);
            if(dataSet.Tables.Count == 0)
                return new DataTable();
            return dataSet.Tables[sheetIndex];
        }



        /// <summary>Convert a single worksheet inside an Excel file into a DataTable.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a single worksheet inside an Excel file into a DataTable.")]
        public DataTable ToDataTable([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Name of the Excel worksheet. If blank will use the first.")]string sheetName, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            var dataSet = this.ToDataSet(filePath,handlerOptions);
            if(dataSet.Tables.Count == 0)
                return new DataTable();
            if(sheetName == "")
                return dataSet.Tables[0];
            int sheetIndex = GetTableIndex(dataSet,sheetName);
            if(sheetIndex == -1)
            {
                this.ErrorMessage = $"Excel file '{filePath}' does not contain worksheet '{sheetName}'";
                return new DataTable();
            }
            return dataSet.Tables[sheetIndex];
        }



        /// <summary>Convert a list of strings into a DataTable.</summary>
        /// <param name="listData">List of strings to convert.</param>
        [Description("Convert a list of strings into a DataTable.")]
        public DataTable ToDataTable([Description("List of strings to convert.")]List<string> listData)
        {
            return this.ToDataTable(listData,this.DefaultOptions);
        }
        /// <summary>Convert a list of strings into a DataTable.</summary>
        /// <param name="listData">List of strings to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a list of strings into a DataTable.")]
        public DataTable ToDataTable([Description("List of strings to convert.")]List<string> listData, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            bool useHeadings = handlerOptions?.UseHeadings ?? this.DefaultOptions.UseHeadings;
            int maxRows = handlerOptions?.MaxRows ?? this.DefaultOptions.MaxRows;
            string tableName = handlerOptions?.DefaultTableName ?? this.DefaultOptions.DefaultTableName;
            string columnName = handlerOptions?.DefaultColumnName ?? this.DefaultOptions.DefaultColumnName;

            if(!this.IsValidWorksheetName(tableName))
                tableName = "text"; // System.Guid.NewGuid().ToString().Replace("-","").Substring(0,31);
            DataTable dataTable = new DataTable(tableName);
            dataTable.TableName = tableName;

            // no header
            dataTable.Columns.Add(columnName);
            // if(useHeadings)
            //     dataTable.Rows.Add(tableName);
            int i = 0;
            foreach(var item in listData)
            {
                dataTable.Rows.Add(item);
                i++;
                if(i==maxRows)
                    break;
            }
            return dataTable;
        }

        /// <summary>Convert a list of integers into a DataTable.</summary>
        /// <param name="listData">List of integers to convert.</param>
        [Description("Convert a list of integers into a DataTable.")]
        public DataTable ToDataTable([Description("List of integers to convert.")]List<int> listData)
        {
            return this.ToDataTable(listData,this.DefaultOptions);
        }
        /// <summary>Convert a list of integers into a DataTable.</summary>
        /// <param name="listData">List of integers to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a list of integers into a DataTable.")]
        public DataTable ToDataTable([Description("List of integers to convert.")]List<int> listData, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            bool useHeadings = handlerOptions?.UseHeadings ?? this.DefaultOptions.UseHeadings;
            int maxRows = handlerOptions?.MaxRows ?? this.DefaultOptions.MaxRows;
            string tableName = handlerOptions?.DefaultTableName ?? this.DefaultOptions.DefaultTableName;
            string columnName = handlerOptions?.DefaultColumnName ?? this.DefaultOptions.DefaultColumnName;

            if(!this.IsValidWorksheetName(tableName))
                tableName = "int"; // System.Guid.NewGuid().ToString().Replace("-","").Substring(0,31);
            DataTable dataTable = new DataTable(tableName);
            dataTable.TableName = tableName;

            // no header
            dataTable.Columns.Add(columnName);
            // if(useHeadings)
            //     dataTable.Rows.Add(columnName);
            int i = 0;
            foreach(var item in listData)
            {
                dataTable.Rows.Add(item);
                i++;
                if(i==maxRows-1)
                    break;
            }
            return dataTable;
        }

        /// <summary>Convert a list of type 'T' into a DataTable.</summary>
        /// <param name="listData">List of type 'T' to convert.</param>
        [Description("Convert a list of type 'T' into a DataTable.")]
        public DataTable ToDataTable<T>([Description("List of type 'T' to convert.")]List<T> listData) where T : new()
        {
            return this.ToDataTable(listData,this.DefaultOptions);
        }
        /// <summary>Convert a list of type 'T' into a DataTable.</summary>
        /// <param name="listData">List of type 'T' to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a list of type 'T' into a DataTable.")]
        public DataTable ToDataTable<T>([Description("List of type 'T' to convert.")]List<T> listData, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new()
        {
            bool useHeadings = handlerOptions?.UseHeadings ?? this.DefaultOptions.UseHeadings;
            int maxRows = handlerOptions?.MaxRows ?? this.DefaultOptions.MaxRows;
            string tableName = handlerOptions?.DefaultTableName ?? this.DefaultOptions.DefaultTableName;

            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties
            var Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var Fields = typeof(T).GetFields(BindingFlags.Public | BindingFlags.Instance);
            // PropertyInfo[] props = t.GetProperties();
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
            }
            foreach (FieldInfo field in Fields)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(field.Name);
            }
            foreach (T item in listData)
            {
                var values = new object[Props.Length+Fields.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    string itemvalue = Props[i]?.GetValue(item, null)?.ToString() ?? "";
                    if(itemvalue == "")
                    {
                        values[i] = "";
                    }
                    else
                    {
                        // string checkitem = this.ShapeText(itemvalue,Props[i].PropertyType,handlerOptions);
                        string checkme = Props[i]?.PropertyType.ToString();
                        string checkitem = (this.ParseExcel(itemvalue,Props[i]?.PropertyType,handlerOptions)).ToString();                        
                        if(checkitem == "" && ( Props[i].GetType() != typeof(string)))
                        {
                            values[i] = "";
                        }
                        else
                        {
                            values[i] = checkitem;
                        }
                    }
                }
                for (int i = Props.Length; i < Props.Length+Fields.Length; i++)
                {
                    //inserting property values to datatable rows
                    // if this is a CLR type, get it as a string, if not, get it as Json
                    string itemvalue = Fields[i-Props.Length]?.GetValue(item)?.ToString() ?? "";
                    if(itemvalue == "")
                    {
                        values[i] = "";
                    }
                    else
                    {
                        // 2023-07-14 SNJW Error where List<T> has non-CLR types #3
                        // The values[] array is all strings - it's just convienient for it to be the object type to add 
                        // directly into the DataTable
                        // however if it's a non-CLR type it won't be Json here
                        Type fieldtype = Fields[i-Props.Length].FieldType;

                        if(!this.IsCLRType(fieldtype))
                        {
                            // 2023-07-14 SNJW if this is a non-CLR type then either set it to JSON or to blank
                            if(handlerOptions?.ComplexToJson ?? this.DefaultOptions.ComplexToJson)
                            {
                                object? nonclr = Fields[i-Props.Length]?.GetValue(item);
                                values[i] = "{}";
                                if(nonclr != null)
                                {
                                    try
                                    {
                                        values[i] = JsonConvert.SerializeObject(nonclr, fieldtype, null);
                                    }
                                    catch
                                    {
                                        
                                    }
                                }
                            }
                            else
                            {
                                values[i] = "";
                            }
                        }
                        else
                        {
                            string checkitem = (this.ParseExcel(itemvalue,Fields[i-Props.Length].FieldType,handlerOptions)).ToString();

                            // string checkitem = this.ShapeText(itemvalue,Fields[i-Props.Length].FieldType,handlerOptions);
                            if(checkitem == "" && ( Fields[i-Props.Length].GetType() != typeof(string)))
                            {
                                values[i] = "";
                            }
                            else
                            {
                                values[i] = checkitem;
                            }
                        }

                    }
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }

        /// <summary>Convert the first column of a DataTable into a list of strings.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        [Description("Convert the first column of a DataTable into a list of strings.")]
        public List<string> ToListDataString([Description("DataTable to convert.")]DataTable dataTable)
        {
            return this.ToListDataString(dataTable,this.DefaultOptions);
        }
        /// <summary>Convert the first column of a DataTable into a list of strings.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert the first column of a DataTable into a list of strings.")]
        public List<string> ToListDataString([Description("DataTable to convert.")]DataTable dataTable, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            List<string> listData = new List<string>();
            if(dataTable.Columns.Count == 0)
            {
                this.ErrorMessage = "No columns in DataTable";
                return listData;
            }
            foreach(System.Data.DataRow dataRow in dataTable.Rows)
                listData.Add((dataRow?.ItemArray?[0] ?? "").ToString());
            return listData;
        }

        /// <summary>Convert the first column of a DataTable into a list of integers.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        [Description("Convert the first column of a DataTable into a list of integers.")]
        public List<Int32> ToListDataInt32([Description("DataTable to convert.")]DataTable dataTable)
        {
            return this.ToListDataInt32(dataTable,this.DefaultOptions);
        }
        /// <summary>Convert the first column of a DataTable into a list of integers.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert the first column of a DataTable into a list of integers.")]
        public List<Int32> ToListDataInt32([Description("DataTable to convert.")]DataTable dataTable, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            if(handlerOptions == null)
                handlerOptions = this.DefaultOptions;
            List<int> listData = new List<int>();
            foreach(System.Data.DataRow dataRow in dataTable.Rows)
            {
                string item = (dataRow?.ItemArray?[0] ?? "").ToString().Trim();
                if(item!="")
                {
                    // ParseExcel
                    item = (string)this.ParseExcel(item,typeof(Int32),handlerOptions);
                    // item = this.ShapeText(item,typeof(Int32),handlerOptions);
                    if(item!="")
                    {
                        int newitem = 0;
                        if(Int32.TryParse(item,out newitem))
                            listData.Add(newitem);
                    }
                }
            }
            return listData;
        }

        /// <summary>Convert a DataTable to a list of type 'T'.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        [Description("Convert a DataTable to list of type 'T'.")]
        public List<T> ToListData<T>([Description("DataTable to convert.")]DataTable dataTable) where T : new()
        {
            return this.ToListData<T>(dataTable,this.DefaultOptions);
        }
        /// <summary>Convert a DataTable to a list of type 'T'.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a DataTable to a list of type 'T'.")]
        public List<T> ToListData<T>([Description("DataTable to convert.")]DataTable dataTable, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new()
        {
            bool useHeadings = handlerOptions?.UseHeadings ?? this.DefaultOptions.UseHeadings;
            List<T> listData = new();
            // Get the properties of the POCO object
            var properties = typeof(T).GetProperties();
            var fields = typeof(T).GetFields();
            string typeName = typeof(T).Name;

            // Iterate through the rows of the DataTable
            foreach (DataRow dataRow in dataTable.Rows)
            {
                T item = new T();
                int propertyTotal = properties.Count();
                int fieldTotal = fields.Count();
                
                for(int i = 0; i < propertyTotal+fieldTotal; i++)
                {
                    bool isProperty = (i < propertyTotal);
                    string destName = "";
                    Type destType = typeof(object);
                    object sourceValue = null;
                    object destValue = null;
                    if(isProperty)
                    {
                        destName = properties[i].Name;
                        destType = properties[i].PropertyType;
                    }
                    else
                    {
                        destName = fields[i-propertyTotal].Name;
                        destType = fields[i-propertyTotal].FieldType;
                    }
                    if(useHeadings)
                    {
                        if (dataTable.Columns.Contains(destName))
                        {
                            sourceValue = dataRow[destName];
                        }
                    }
                    else
                    {
                        if (i < dataTable.Columns.Count)
                        {
                            // Set the value of the property
                            // Try to parse input
                            sourceValue = dataRow[i];
                        }
                    }
                    if(sourceValue != null)
                    {
                        // if a non-CLR type, check whether it's a JSON object
                        if(!this.IsCLRType(destType) && (handlerOptions?.ComplexToJson ?? this.DefaultOptions.ComplexToJson))
                        {
                            destValue = this.ParseExcelComplex(sourceValue, destType, handlerOptions);
                        }
                        else
                        {
                            destValue = this.ParseExcel(sourceValue, destType, handlerOptions);
                        }

                        if(isProperty)
                        {
                            properties[i].SetValue(item, destValue);
                        }
                        else
                        {
                            fields[i-propertyTotal].SetValue(item, destValue);
                        }
                    }
                }

                // Add the POCO object to the result
                listData.Add(item);
            }
            return listData;
        }
  
        /// <summary>Parses an Excel data value and returns output of the specified type.</summary>
        /// <param name="sourceValue">Source value from Excel.</param>
        /// <param name="destType">Destination type.</param>
        [Description("Parses an Excel data value and returns output of the specified type.")]
        public object ParseExcel([Description("Source value from Excel.")]object sourceValue, [Description("Destination type.")]Type destType)
        {
            return this.ParseExcel(sourceValue,destType,this.DefaultOptions);
        }
        /// <summary>Parses an Excel data value and returns output of the specified type.</summary>
        /// <param name="sourceValue">Source value from Excel.</param>
        /// <param name="destType">Destination type.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Parses an Excel data value and returns output of the specified type.")]
        public object ParseExcel([Description("Source value from Excel.")]object sourceValue, [Description("Destination type.")]Type destType, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            if(sourceValue==null)
            {
                if(handlerOptions?.NullValueAction == ExcelToDataOptions.NullValue.None)
                {
                    return sourceValue;
                }
                else if(handlerOptions?.NullValueAction == ExcelToDataOptions.NullValue.Blank)
                {
                    if(destType == typeof(bool))
                    {
                        return (object)handlerOptions?.BlankValues?.BooleanDefaultValue;
                    }
                    else if(destType == typeof(string))
                    {
                        return (object)handlerOptions?.BlankValues?.StringDefaultValue;
                    }
                    else if(destType == typeof(DateTime?))
                    {
                        return (object)handlerOptions?.BlankValues?.NullableDateTimeDefaultValue;
                    }
                    else if(destType == typeof(DateTime))
                    {
                        return (object)handlerOptions?.BlankValues?.DateTimeDefaultValue;
                    }
                    else if(destType == typeof(int))
                    {
                        return (object)handlerOptions?.BlankValues?.Int32DefaultValue;
                    }
                    else if(destType == typeof(Single))
                    {
                        return (object)handlerOptions?.BlankValues?.SingleDefaultValue;
                    }
                    else if(destType == typeof(double))
                    {
                        return (object)handlerOptions?.BlankValues?.DoubleDefaultValue;
                    }
                    else if(destType == typeof(decimal))
                    {
                        return (object)handlerOptions?.BlankValues?.DecimalDefaultValue;
                    }
                    else if(destType == typeof(byte))
                    {
                        return (object)handlerOptions?.BlankValues?.ByteDefaultValue;
                    }
                    else if(destType == typeof(short))
                    {
                        return (object)handlerOptions?.BlankValues?.ShortDefaultValue;
                    }
                    else if(destType == typeof(long))
                    {
                        return (object)handlerOptions?.BlankValues?.LongDefaultValue;
                    }
                    else if(destType == typeof(TimeSpan))
                    {
                        return (object)handlerOptions?.BlankValues?.NullableTimeSpanDefaultValue;
                    }
                    else if(destType == typeof(Guid))
                    {
                        return (object)handlerOptions?.BlankValues?.NullableGuidDefaultValue;
                    }
                    else
                    {
                        return null;
                    }
                }
                else if(handlerOptions?.NullValueAction == ExcelToDataOptions.NullValue.Default)
                {

                    if(destType.IsValueType)
                    {
                        return Activator.CreateInstance(destType);
                    }
                    else
                    {
                        return null;
                    }
                }
                return sourceValue;
            }
            else
            {
                object destValue = null;
                if(destType == typeof(DateTime))
                {
                    destValue = this.ParseExcelDateTime(sourceValue, handlerOptions);
                }
                else if(destType == typeof(DateTime?))
                {
                    destValue = this.ParseExcelDateTimeNullable(sourceValue, handlerOptions);
                }
                else if(destType == typeof(bool))
                {
                    destValue = this.ParseExcelBool(sourceValue, handlerOptions);
                }
                else if(destType == typeof(bool?))
                {
                    destValue = this.ParseExcelBoolNullable(sourceValue, handlerOptions);
                }
                else if(!this.IsCLRType(destType))
                {
                    // 2023-07-14 SNJW if this is a non-CLR type then either set it to JSON or to blank
                    if(handlerOptions?.ComplexToJson ?? this.DefaultOptions.ComplexToJson)
                    {
                        // if(handlerOptions?.ComplexNullable ?? this.DefaultOptions.ComplexNullable)
                        //     destValue = this.ParseExcelComplexNullable(sourceValue, destType, handlerOptions);
                        // string checkJson = JsonConvert.SerializeObject(sourceValue, destType, null);
                        object checkValue = this.ParseExcelComplex(sourceValue, destType, handlerOptions);
                        destValue = JsonConvert.SerializeObject(checkValue, destType, null);
                        // destValue = this.ParseExcelComplex(sourceValue, destType, handlerOptions);
                        // destValue = this.ParseExcelJson(checkJson, destType, handlerOptions);
                    }

                    // TO DO put code that instantiates a default object if the user needs this and the object cannot be null
                }
                else
                {
                    try
                    {
                        destValue = (object)Convert.ChangeType(sourceValue, destType);
                    }
                    catch (FormatException)
                    {
                        // silently fail here
                    }
                }
                return destValue;
            }
        }


        /// <summary>Parses an Excel data value formatted in JSON and return it as the specified non-CLR type.</summary>
        /// <param name="sourceValue">Source value from Excel.</param>
        /// <param name="destType">Class name.</param>
        [Description("Parses an Excel data value that is formatted in JSON and returns it if well-formed.")]
        public object ParseExcelComplex([Description("Source value from Excel.")]object sourceValue, [Description("Target class/type.")]Type destType)
        {
            return this.ParseExcelComplex(sourceValue,destType,this.DefaultOptions);
        }

        /// <summary>Parses an Excel data value formatted in JSON and return it as the specified non-CLR type.</summary>
        /// <param name="sourceValue">Source value from Excel.</param>
        /// <param name="destType">Class name.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Parses an Excel data value that is formatted in JSON and returns it if well-formed.")]
        public object ParseExcelComplex([Description("Source value from Excel.")]object sourceValue, [Description("Target class/type.")]Type destType,[Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            try
            {
                object check = JsonConvert.DeserializeObject(sourceValue.ToString(), destType);
                return check;
            }
            catch
            {
            }
            try
            {
                Assembly assembly = destType.Assembly;
                object check = (object)assembly.CreateInstance(destType.FullName);
                return check;
            }
            catch
            {
            }
            return new Object();
        }

        /// <summary>Parses an Excel data value that is formatted in JSON and returns it if well-formed.</summary>
        /// <param name="sourceValue">Source value from Excel.</param>
        /// <param name="destType">Class name.</param>
        [Description("Parses an Excel data value that is formatted in JSON and returns it if well-formed.")]
        public string ParseExcelJson([Description("Source value from Excel.")]object sourceValue, [Description("Target class/type.")]Type destType)
        {
            return this.ParseExcelJson(sourceValue,destType,this.DefaultOptions);
        }

        /// <summary>Parses an Excel data value that is formatted in JSON and returns it if well-formed.</summary>
        /// <param name="sourceValue">Source value from Excel.</param>
        /// <param name="destType">Class name.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Parses an Excel data value that is formatted in JSON and returns it if well-formed.")]
        public string ParseExcelJson([Description("Source value from Excel.")]object sourceValue, [Description("Target class/type.")]Type destType,[Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            if(sourceValue == null)
                return "{}";
            string json = (sourceValue ?? "").ToString();
            if(!this.IsValidJson(json))
                return "{}";
            string output = "{}";
            try
            {
                object check = JsonConvert.DeserializeObject(json, destType);
                output = JsonConvert.SerializeObject(check);
            }
            catch
            {
                output = "{}";
            }
            return output;
        }

        private bool IsValidJson(string strInput)
        {
            if (string.IsNullOrWhiteSpace(strInput))
                return false;
            strInput = strInput.Trim();
            bool success = false;
            if ((strInput.StartsWith("{") && strInput.EndsWith("}")) || //For object
                (strInput.StartsWith("[") && strInput.EndsWith("]"))) //For array
            {
                try
                {
                    var obj = JToken.Parse(strInput);
                    success = true;
                }
                catch
                {
                    success = false;
                }
            }
            return success;
        }

        public bool IsCLRType(Type type)
        {
            var fullname = type.Assembly.FullName;
            return (fullname??"").StartsWith("System.Private.CoreLib");
        }        
        /// <summary>Parses an Excel data value and returns its true/false/null status.</summary>
        /// <param name="sourceValue">Source value from Excel.</param>
        [Description("Parses an Excel data value and returns its true/false/null status.")]
        public bool? ParseExcelBoolNullable([Description("Source value from Excel.")]object sourceValue)
        {
            return this.ParseExcelBoolNullable(sourceValue,this.DefaultOptions);
        }
        /// <summary>Parses an Excel data value and returns its true/false/null status.</summary>
        /// <param name="sourceValue">Source value from Excel.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Parses an Excel data value and returns its true/false/null status.")]
        public bool? ParseExcelBoolNullable([Description("Source value from Excel.")]object sourceValue, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            if(sourceValue == null)
                return null;
            if(sourceValue.ToString().ToLower() == "null" || sourceValue.ToString().ToLower() == "\"null\"")
                return null;
            return (bool?)ParseExcelBool(sourceValue,handlerOptions);
        }

        /// <summary>Parses an Excel data value and returns its true/false status.</summary>
        /// <param name="sourceValue">Source value from Excel.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Parses an Excel data value and returns its true/false status.")]
        public bool ParseExcelBool([Description("Source date/time value from Excel.")]object sourceValue,[Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            Type sourceType = sourceValue.GetType();
            string checkstring = sourceValue?.ToString().Trim();
            if(checkstring == ""
                || checkstring == "0"
                || checkstring == "False"
                || checkstring == "false"
                || checkstring == "f"
                || checkstring == "F")
                return false;
            checkstring = checkstring.ToUpper();
            if(checkstring == "FALSE"
                || checkstring == "NONE"
                || checkstring == "NIL"
                || checkstring == "NO"
                || checkstring == "ZERO"
                || checkstring == "\"FALSE\""
                || checkstring == "\"0\""
                || checkstring == "\"F\""
                )
                return false;
            return true;
        }
        /// <summary>Parses an Excel data value and returns its true/false status.</summary>
        /// <param name="sourceValue">Source value from Excel.</param>
        [Description("Parses an Excel data value and returns its true/false status.")]
        public bool ParseExcelBool([Description("Source date/time value from Excel.")]object sourceValue)
        {
            return this.ParseExcelBool(sourceValue,this.DefaultOptions);
        }

        /// <summary>Parses an Excel date/time value and returns its nullable CLR DateTime value.</summary>
        /// <param name="sourceValue">Source date/time value from Excel.</param>
        [Description("Parses an Excel date/time value and returns its nullable CLR DateTime value.")]
        public DateTime? ParseExcelDateTimeNullable([Description("Source date/time value from Excel.")]object sourceValue)
        {
            return this.ParseExcelDateTimeNullable(sourceValue,this.DefaultOptions);
        }
        /// <summary>Parses an Excel date/time value and returns its nullable CLR DateTime value.</summary>
        /// <param name="sourceValue">Source date/time value from Excel.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Parses an Excel date/time value and returns its nullable CLR DateTime value.")]
        public DateTime? ParseExcelDateTimeNullable([Description("Source date/time value from Excel.")]object sourceValue, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            Type sourceType = sourceValue.GetType();
            int checkint = -1;
            decimal checkdecimal = -1.00M;
            bool isInt32 = false;
            bool isDecimal = false;

            if(sourceType==typeof(string))
            {
                // this can be "30/04/2023 4:54:50 PM"
                // or "30/04/2023"
                // or "48134"
                // or "48134.6541"
                string checkstring = sourceValue?.ToString().Trim();

                if((handlerOptions?.DateTimeCheckExcelSerial ?? false) == true)
                {
                    //string lhs = checkstring;
                    //string rhs = "";
                    if(!Int32.TryParse(checkstring,out checkint))
                        checkint = -1;
                    if(checkint <= 0 || checkint > 2885415)
                    {
                        checkint = -1;
                    }
                    else
                    {
                        isInt32 = true;
                    }

                    if(!isInt32)
                    {
                        if(!Decimal.TryParse(checkstring,out checkdecimal))
                            checkdecimal = -1.00M;
                        if(checkdecimal <= 0.00M || checkdecimal > 2885415.00M)
                        {
                            checkdecimal = -1.00M;
                        }
                        else
                        {
                            isDecimal = true;
                        }
                    }
                }

                if(!isInt32 && !isDecimal && (handlerOptions?.DateTimeCheckSourceFormat ?? false) == true)
                {
                    DateTime dateValue;
                    if (DateTime.TryParseExact(checkstring, handlerOptions?.SourceDateTimeFormat, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateValue))
                        return (DateTime?)dateValue;
                    if (DateTime.TryParseExact(checkstring, handlerOptions?.SourceDateFormat, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateValue))
                        return (DateTime?)dateValue;
                }

                if(!isInt32 && !isDecimal && (handlerOptions?.DateTimeCheckFixed ?? false) == true)
                {
                    DateTime dateValue;
                    if (DateTime.TryParse(checkstring, out dateValue))
                    {
                        return (DateTime?)dateValue;
                    }
                    foreach(var cultureinfo in handlerOptions.DateTimeCheckCultures)
                    {
                        foreach(string dateformat in handlerOptions.DateTimeCheckFormats)
                        {
                            if (DateTime.TryParseExact(checkstring, dateformat, cultureinfo, System.Globalization.DateTimeStyles.AllowWhiteSpaces, out dateValue))
                            {
                                return (DateTime?)dateValue;
                            }
                        }
                    }
                }
            }
            else if(sourceType==typeof(int))
            {
                checkint = (int)sourceValue;
            }
            else if(sourceType==typeof(Single) 
                || sourceType==typeof(float)
                || sourceType==typeof(decimal)
                || sourceType==typeof(Double))
            {
                checkdecimal = (decimal)sourceValue;
            }

            if(isInt32 && checkint >= 0 && checkint < 2885415)
            {
                return (DateTime?)(DateTime.FromOADate((double)checkint));
            }
            else if(isDecimal && checkdecimal >= 0.00M && checkdecimal < 2885415.00M)
            {
                return (DateTime?)(DateTime.FromOADate((double)checkdecimal));
            }

            return null;
        }

        

        /// <summary>Performs a case-insensitive check of table names in a DataSet.  Returns -1 if none found.</summary>
        /// <param name="dataSet">Source data from Excel.</param>
        /// <param name="checkName">Worksheet name to check.</param>
        [Description("Performs a case-insensitive check of table names in a DataSet.  Returns -1 if none found.")]
        private int GetTableIndex([Description("DataSet to search.")]DataSet dataSet, [Description("Name of the DataTable.")]string checkName)
        {
            if(checkName=="")
                return -1;
            checkName = checkName.Trim().ToLower();
            for(int i = 0; i < dataSet.Tables.Count; i++)
            {
                // 2023-07-15 SNJW MS Excel worksheet names are case-insensitve
                // if(dataSet.Tables[i]?.TableName == tableName)
                if(dataSet.Tables[i]?.TableName.Trim().ToLower() == checkName)
                    return i;
            }
            return -1;
        }


        /// <summary>Parses an Excel date/time value and returns its CLR DateTime value.</summary>
        /// <param name="sourceValue">Source date/time value from Excel.</param>
        [Description("Parses an Excel date/time value and attempts to determine the CLR nullable DateTime value.")]
        public DateTime ParseExcelDateTime([Description("Source date/time value from Excel.")]object sourceValue)
        {
            return this.ParseExcelDateTime(sourceValue,this.DefaultOptions);
        }
        /// <summary>Parses an Excel date/time value and returns its CLR DateTime value.</summary>
        /// <param name="sourceValue">Source date/time value from Excel.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Parses an Excel date/time value and attempts to determine the CLR nullable DateTime value.")]
        public DateTime ParseExcelDateTime([Description("Source date/time value from Excel.")]object sourceValue,[Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            DateTime? checkdate = this.ParseExcelDateTimeNullable(sourceValue,handlerOptions);
            if(checkdate != null)
            {
                return (DateTime)checkdate;
            }
            else if(handlerOptions?.DateTimeDefaultAction == ExcelToDataOptions.DateTimeDefault.Lowest)
            {
                return DateTime.MinValue;
            }
            else if(handlerOptions?.DateTimeDefaultAction == ExcelToDataOptions.DateTimeDefault.Set1900)
            {
                return new DateTime(1900,1,1);
            }
            else if(handlerOptions?.DateTimeDefaultAction == ExcelToDataOptions.DateTimeDefault.Set1970)
            {
                return new DateTime(1970,1,1);
            }
            else if(handlerOptions?.DateTimeDefaultAction == ExcelToDataOptions.DateTimeDefault.Set1980)
            {
                return new DateTime(1980,1,1);
            }
            else if(handlerOptions?.DateTimeDefaultAction == ExcelToDataOptions.DateTimeDefault.Set1990)
            {
                return new DateTime(1990,1,1);
            }
            else if(handlerOptions?.DateTimeDefaultAction == ExcelToDataOptions.DateTimeDefault.Set2000)
            {
                return new DateTime(2000,1,1);
            }
            return System.DateTime.Now;
        }


#region csv-handling


        /// <summary>Convert a list of type 'T' into columns in a CSV file.</summary>
        /// <param name="listData">List of type 'T' to convert.</param>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        [Description("Convert a list of type 'T' into columns in a CSV file.")]
        public bool ToCsvFile<T>([Description("List of type 'T' convert.")]List<T> listData, [Description("Full path and name of the CSV file.")]string filePath) where T : new() 
        {
            return this.ToCsvFile<T>(listData,filePath,this.DefaultOptions);
        }
        /// <summary>Convert a list of type 'T' into columns in a CSV file.</summary>
        /// <param name="listData">List of type 'T' to convert.</param>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a list of type 'T' into columns in a CSV file.")]
        public bool ToCsvFile<T>([Description("List of type 'T' convert.")]List<T> listData, [Description("Full path and name of the CSV file.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new() 
        {
            ErrorMessage = "";
            var dataTable = this.ToDataTable<T>(listData, handlerOptions);
            string csvText = this.ToCsvText(dataTable, handlerOptions);
            if(ErrorMessage!="")
                return false;
            return this.ToCsvFile(csvText, filePath, handlerOptions);
        }

        /// <summary>Convert a list of strings into one column in a CSV file.</summary>
        /// <param name="listData">List of strings to convert.</param>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        [Description("Convert a list of strings into one column in a CSV file.")]
        public bool ToCsvFile([Description("List of strings to convert.")]List<string> listData, [Description("Full path and name of the CSV file.")]string filePath)
        {
            return this.ToCsvFile(listData, filePath, this.DefaultOptions);
        }
        /// <summary>Convert a list of strings into one column in a CSV file.</summary>
        /// <param name="listData">List of strings to convert.</param>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a list of strings into one column in a CSV file.")]
        public bool ToCsvFile([Description("List of strings to convert.")]List<string> listData, [Description("Full path and name of the CSV file.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            ErrorMessage = "";
            DataTable dataTable = this.ToDataTable(listData, handlerOptions);
            if(ErrorMessage!="")
                return false;
            string csvText = this.ToCsvText(dataTable, handlerOptions);
            if(ErrorMessage!="")
                return false;
            return this.ToCsvFile(csvText, filePath, handlerOptions);
        }

        /// <summary>Convert a list of integers into one column in a CSV file.</summary>
        /// <param name="listData">List of integers to convert.</param>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        [Description("Convert a list of integers into one column in a CSV file.")]
        public bool ToCsvFile([Description("List of integers to convert.")]List<int> listData, [Description("Full path and name of the CSV file.")]string filePath)
        {
            return this.ToCsvFile(listData,"",this.DefaultOptions);
        }
        /// <summary>Convert a list of integers into one column in a CSV file.</summary>
        /// <param name="listData">List of integers to convert.</param>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a list of integers into one column in a CSV file.")]
        public bool ToCsvFile([Description("List of integers to convert.")]List<int> listData, [Description("Full path and name of the CSV file.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            ErrorMessage = "";
            DataTable dataTable = this.ToDataTable(listData, handlerOptions);
            if(ErrorMessage!="")
                return false;
            string csvText = this.ToCsvText(dataTable, handlerOptions);
            if(ErrorMessage!="")
                return false;
            return this.ToCsvFile(csvText, filePath, handlerOptions);
        }

        /// <summary>Convert a DataSet into a CSV file.</summary>
        /// <param name="dataSet">DataSet containing DataTables to convert.</param>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        [Description("Convert a DataSet into a CSV file.")]
        public bool ToCsvFile([Description("DataSet containing DataTables to convert.")]DataSet dataSet, [Description("Full path and name of the CSV file.")]string filePath)
        {
            return this.ToCsvFile(dataSet,filePath,this.DefaultOptions);
        }
        /// <summary>Convert a DataSet into a CSV file.</summary>
        /// <param name="dataSet">DataSet containing DataTables to convert.</param>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a DataSet into a CSV file.")]
        public bool ToCsvFile([Description("DataSet containing DataTables to convert.")]DataSet dataSet, [Description("Full path and name of the CSV file.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            ErrorMessage = "";
            string csvText = this.ToCsvText(dataSet,handlerOptions)[0];
            if(ErrorMessage!="")
                return false;
            return this.ToCsvFile(csvText, filePath,handlerOptions);
        }


        /// <summary>Convert a DataSet into multiple CSV files.</summary>
        /// <param name="dataSet">DataSet containing DataTables to convert.</param>
        /// <param name="directoryPath">Full path to the output directory.</param>
        [Description("Convert a DataSet into a CSV file.")]
        public bool ToCsvFiles([Description("DataSet containing DataTables to convert.")]DataSet dataSet, [Description("Full path to output directory.")]string directoryPath)
        {
            return this.ToCsvFile(dataSet,directoryPath,this.DefaultOptions);
        }
        /// <summary>Convert a DataSet into multiple CSV files.</summary>
        /// <param name="dataSet">DataSet containing DataTables to convert.</param>
        /// <param name="directoryPath">Full path to the output directory.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a DataSet into multiple CSV files.")]
        public bool ToCsvFiles([Description("DataSet containing DataTables to convert.")]DataSet dataSet, [Description("Full path to output directory.")]string directoryPath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            ErrorMessage = "";
            var csvFiles = this.ToCsvText(dataSet,handlerOptions);
            if(csvFiles.Count < dataSet.Tables.Count)
                if(ErrorMessage=="")
                    ErrorMessage = $"Could only create {csvFiles.Count} files from {dataSet.Tables.Count} files"; 
            if(ErrorMessage!="")
                return false;
            bool saveResult = true;
            for (int i = 0; i < dataSet.Tables.Count; i++)
            {
                string filePath = Path.Combine(directoryPath,dataSet.Tables[i].TableName+".csv");
                if(!this.ToCsvFile(csvFiles[i],filePath,handlerOptions))
                    saveResult = false;
            }
            if(ErrorMessage!="")
                return false;
            return saveResult;
        }

        /// <summary>Convert a DataTable into a CSV file.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        [Description("Convert a DataTable into a CSV file.")]
        public bool ToCsvFile([Description("DataTable to convert.")]DataTable dataTable, [Description("Full path and name of the CSV file.")]string filePath)
        {
            return this.ToCsvFile(dataTable,filePath,this.DefaultOptions);
        }
        /// <summary>Convert a DataTable into a CSV file.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a DataSet into a CSV file.")]
        public bool ToCsvFile([Description("DataTable to convert.")]DataTable dataTable, [Description("Full path and name of the CSV file.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable.Copy());
            return this.ToCsvFile(dataSet,filePath,handlerOptions);
        }

        /// <summary>Save an in-memory CSV text file to a local file.</summary>
        /// <param name="csvText">CSV text data to save as a file.</param>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        [Description("Save an in-memory CSV text file to a local file.")]
        public bool ToCsvFile([Description("CSV text data to save as a file.")]string csvText, [Description("Full path and name of the CSV file.")]string filePath)
        {
            return this.ToCsvFile(csvText,filePath,this.DefaultOptions);
        }
        /// <summary>Save an in-memory CSV text file to a local file.</summary>
        /// <param name="csvText">CSV text data to save as a file.</param>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Save in-memory CSV text to a local file.")]
        public bool ToCsvFile([Description("CSV text data to save as a file.")]string csvText, [Description("Full path and name of the CSV file.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            ErrorMessage = "";
            if(ErrorMessage!="")
                return false;
            try
            {
                File.WriteAllText(filePath, csvText);
            }
            catch (IOException ex)
            {
                ErrorMessage = $"Error creating CSV file: '{ex.InnerException}'";
            }
            return ErrorMessage=="";
        }

        /// <summary>Load a local CSV file to in-memory string data.</summary>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        [Description("Load a local CSV file to in-memory string data.")]
        public string ToCsvText([Description("Full path and name of the CSV file.")]string filePath)
        {
            return this.ToCsvText(filePath,this.DefaultOptions);
        }
        /// <summary>Load a local CSV file to in-memory string data.</summary>
        /// <param name="filePath">Full path and name of the CSV file.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Load a local CSV file to in-memory string data.")]
        public string ToCsvText([Description("Full path and name of the CSV file.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            string result ="";
            ErrorMessage = "";
            try
            {
                result = File.ReadAllText(filePath);
            }
            catch (IOException ex)
            {
                ErrorMessage = $"Error loading CSV file: '{ex.InnerException}'";
            }
            return result;
        }

        /// <summary>Convert a list of type 'T' into an in-memory string CSV file.</summary>
        /// <param name="listData">List of type 'T' to copy into an in-memory CSV file.</param>
        [Description("Convert a list of type 'T' into an in-memory string CSV file.")]
        public string ToCsvText<T>([Description("List of type 'T' to copy into an in-memory CSV file.")]List<T> listData) where T : new()
        {
            return this.ToCsvText<T>(listData,this.DefaultOptions);
        }
        /// <summary>Convert a list of type 'T' into an in-memory string CSV file.</summary>
        /// <param name="listData">List of type 'T' to copy into an in-memory CSV file.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a list of type 'T' into an in-memory string CSV file.")]
        public string ToCsvText<T>([Description("List of type 'T' to copy into an in-memory CSV file.")]List<T> listData, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new()
        {
            var dataTable = this.ToDataTable<T>(listData,handlerOptions);
            return this.ToCsvText(dataTable,handlerOptions);
        }

        /// <summary>Convert a DataTable into an in-memory string CSV file.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        [Description("Convert a DataTable into an in-memory string CSV file.")]
        public string ToCsvText([Description("DataTable to convert.")]DataTable dataTable)
        {
            return this.ToCsvText(dataTable,this.DefaultOptions);
        }
        /// <summary>Convert a DataTable into an in-memory string CSV file.</summary>
        /// <param name="dataTable">DataTable to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>        
        [Description("Convert a DataTable into an in-memory string CSV file.")]
        public string ToCsvText([Description("DataTable to convert.")]DataTable dataTable, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);
            return this.ToCsvText(dataSet,handlerOptions)[0];
        }

        /// <summary>Convert a DataSet into an in-memory string CSV file.</summary>
        /// <param name="dataSet">DataSet containing DataTables to convert.</param>
        [Description("Convert a DataSet into an in-memory string CSV file.")]
        public List<string> ToCsvText([Description("DataSet containing DataTables to convert.")]DataSet dataSet)
        {
            return this.ToCsvText(dataSet,this.DefaultOptions);
        }
        /// <summary>Convert a DataSet into an in-memory string CSV file.</summary>
        /// <param name="dataSet">DataSet containing DataTables to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>        
        [Description("Convert a DataSet into an in-memory string CSV file.")]
        public List<string> ToCsvText([Description("DataSet containing DataTables to convert.")]DataSet dataSet, ExcelToDataOptions handlerOptions)
        {
            bool useHeadings = handlerOptions?.UseHeadings ?? this.DefaultOptions.UseHeadings;
            int maxRows = handlerOptions?.MaxRows ?? this.DefaultOptions.MaxRows;
            List<string> columnsToDateTime = handlerOptions?.ColumnsToDateTime ?? new List<string>();
            List<string> columnsToNumber = handlerOptions?.ColumnsToNumber ?? new List<string>();

            string outputDateFormat = handlerOptions?.OutputDateFormat ?? "yyyy-MM-dd";
            string outputDateTimeFormat = handlerOptions?.OutputDateTimeFormat ?? "yyyy-MM-dd hh:mm:ss";
            bool outputDateOnly = handlerOptions?.OutputDateOnly ?? false;

            // 1.0.4 2024-02-04 SNJW add CSV-specific options
            bool csvWrapAll = handlerOptions?.CsvWrapAll ?? false;
            string csvFormat = handlerOptions?.CsvFormat ?? "UTF-8 (Comma delimited)";
            string csvNewLine = handlerOptions?.CsvNewLine ?? "\r\n";
            // CSV UTF-8 (Comma delimited)
            // CSV (Comma delimited)
            // CSV (Macintosh)
            // CSV (MS-DOS)
            // UTF-16 Unicode Text (.txt)
            List<string> output = new();

            bool addDateCellStyle = false;
            bool addNumberCellStyle = false;
            if(columnsToDateTime.Count > 0 || columnsToNumber.Count > 0)
            {
                foreach(System.Data.DataTable table in dataSet.Tables)
                {
                    foreach(System.Data.DataColumn column in table.Columns)
                    {
                        if(columnsToDateTime.Contains(table.TableName+'.'+column.ColumnName) || columnsToDateTime.Contains(column.ColumnName))
                        {
                            addDateCellStyle = true;
                        }
                        if(columnsToNumber.Contains(table.TableName+'.'+column.ColumnName) || columnsToNumber.Contains(column.ColumnName))
                        {
                            addNumberCellStyle = true;
                        }

                        if(addDateCellStyle && addNumberCellStyle)
                            break;
                    }
                }
            }

            ErrorMessage = "";

            int tableid = 0;

            foreach(System.Data.DataTable dataTable in dataSet.Tables)
            {
                // Append a new worksheet and associate it with the workbook
                System.Text.StringBuilder sb = new();
                tableid++;
                string tableName = (dataTable.TableName ?? "");
                if(!this.IsValidWorksheetName(tableName))
                {
                    tableName = "Sheet"+tableid.ToString();
                    if(dataSet.Tables.Cast<DataTable>().Any(table => table.TableName == tableName))
                        tableName = System.Guid.NewGuid().ToString().Replace("-","").Substring(0,31);
                }

                // Add the header row
                if(useHeadings)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        string columnName = dataTable.Columns[i].ColumnName;
                        if(csvWrapAll || columnName.Contains(','))
                            sb.Append('"');
                        sb.Append(columnName);
                        if(csvWrapAll || columnName.Contains(','))
                            sb.Append('"');
                        if(i==dataTable.Columns.Count-1)
                        {
                            sb.Append(csvNewLine);
                        }
                        else
                        {
                            sb.Append(',');
                        }
                    }
                }

                // Add the data rows
                int rowcount = 0;
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        string cellData = (dataTable.Rows[i].ItemArray[j] ?? "").ToString();
                        if(csvWrapAll || cellData.Contains(','))
                            sb.Append('"');
                        if(cellData.Contains('"') && ( csvWrapAll || cellData.Contains(',') ))
                        {
                            sb.Append(cellData.Replace("\"","\\\""));
                        }
                        else if(cellData.Contains('\r') || cellData.Contains('\n'))
                        {
                            sb.Append(cellData.Replace("\n","\\n").Replace("\r","\\r"));
                        }
                        else
                        {
                            sb.Append(cellData);
                        }

                        if(csvWrapAll || cellData.Contains(','))
                            sb.Append('"');
                        if(j==dataTable.Columns.Count-1)
                        {
                            sb.Append(csvNewLine);
                        }
                        else
                        {
                            sb.Append(',');
                        }

                    }
                    if(rowcount==maxRows)
                        break;
                    rowcount++;
                }
                output.Add(sb.ToString());
            }

            return output;
        }


        /// <summary>Convert in-memory CSV text into a DataSet.</summary>
        /// <param name="csvText">In-memory CSV text to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert an in-memory binary Excel file into a DataSet.")]
        public DataSet CsvTextToDataSet([Description("In-memory CSV text to convert.")]string csvText, [Description("Optional settings.")]ExcelToDataOptions handlerOptions, [Description("DataTable name.")]string tableName)
        {
            bool useHeadings = handlerOptions?.UseHeadings ?? this.DefaultOptions.UseHeadings;
            int maxRows = handlerOptions?.MaxRows ?? this.DefaultOptions.MaxRows;
            List<string> columnsToDateTime = handlerOptions?.ColumnsToDateTime ?? new List<string>();
            List<string> columnsToNumber = handlerOptions?.ColumnsToNumber ?? new List<string>();

            string outputDateFormat = handlerOptions?.OutputDateFormat ?? "yyyy-MM-dd";
            string outputDateTimeFormat = handlerOptions?.OutputDateTimeFormat ?? "yyyy-MM-dd hh:mm:ss";
            bool outputDateOnly = handlerOptions?.OutputDateOnly ?? false;

            // 1.0.4 2024-02-04 SNJW add CSV-specific options
            bool csvWrapAll = handlerOptions?.CsvWrapAll ?? false;
            string csvFormat = handlerOptions?.CsvFormat ?? "UTF-8 (Comma delimited)";
            string csvNewLine = handlerOptions?.CsvNewLine ?? "\r\n";

            DataSet output = new();
            DataTable table = new();
            if(!this.IsValidWorksheetName(tableName))
                tableName = "Sheet1";
            table.TableName=tableName;

            bool addDateCellStyle = false;
            bool addNumberCellStyle = false;

            // split the text into lines
            string[] lines = csvText.Split(csvNewLine);
            List<string> toprow = this.GetCsvFields(lines[0], csvWrapAll);
            if(toprow.Count==0)
                return output;
            int startLine = 0;
            if(useHeadings)
                startLine = 1;

            for (int i = 0; i < toprow.Count; i++)
            {
                System.Data.DataColumn newColumn = new ();
                if(useHeadings)
                {
                    newColumn.ColumnName = this.SanitiseFieldName(toprow[i]);
                }
                else
                {
                    newColumn.ColumnName = "Column" + i.ToString();
                }
                table.Columns.Add(newColumn);
            }

            for (int i = startLine; i < lines.Length; i++)
            {
                System.Data.DataRow newRow = table.NewRow();
                var rowItems = this.GetCsvFields(lines[i], csvWrapAll);
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    // 
                    if(j < rowItems.Count)
                    {
                        newRow[j] = rowItems[j];
                    }
                    else
                    {
                        newRow[j] = "";
                    }
                }
                table.Rows.Add(newRow);
            }
            output.Tables.Add(table);
            return output;
        }

        private List<string> GetCsvFields(string lineText, bool csvWrapAll)
        {
            // please draft a method in C# called GetCsvFields that takes parameters lineText as a string and csvWrapAll as a bool and returns a List<string>.
            // this method iterates through each character in lineText and if csvWrapAll is false then returns one item in output for each comma-separated value except where wrapped with " characters
            // if csvWrapAll is true then only insert fields into the output which are wrapped in " characters 

            var fields = new List<string>();
            bool insideQuotes = false;
            var currentField = "";

            for (int i = 0; i < lineText.Length; i++)
            {
                char c = lineText[i];
                
                // Toggle insideQuotes flag when a quote is encountered
                if (c == '"')
                {
                    insideQuotes = !insideQuotes;
                    
                    // If csvWrapAll is true, do not treat quotes as part of the field value
                    if (csvWrapAll)
                        continue;
                }

                // Handle field separation
                if (c == ',' && !insideQuotes)
                {
                    if (!csvWrapAll || (csvWrapAll && currentField.StartsWith("\"") && currentField.EndsWith("\"")))
                    {
                        // Remove wrapping quotes if present
                        fields.Add(csvWrapAll ? currentField.Substring(1, currentField.Length - 2) : currentField);
                    }
                    currentField = "";
                }
                else
                {
                    currentField += c;
                }
            }

            // Add the last field if not empty
            if (!string.IsNullOrEmpty(currentField))
            {
                if (!csvWrapAll || (csvWrapAll && currentField.StartsWith("\"") && currentField.EndsWith("\"")))
                {
                    // Remove wrapping quotes if present and needed
                    fields.Add(csvWrapAll ? currentField.Substring(1, currentField.Length - 2) : currentField);
                }
            }

            return fields;
        }
        


        #endregion csv-handling



    }
}