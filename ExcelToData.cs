using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;
using System.ComponentModel;
using System.Globalization;

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

        public void SetDefaults(ExcelToDataOptions handlerOptions)
        {
            this.DefaultOptions = handlerOptions;
        }
        public ExcelToDataOptions GetDefaults()
        {
            return this.DefaultOptions;
        }
        private ExcelToDataOptions BaseDefaults()
        {
            return new ExcelToDataOptions()
            {
                MaxRows = 65535,
                UseHeadings = true,
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
                DateTimeCheckCultures = new List<CultureInfo>{ CultureInfo.GetCultureInfo("en-AU")},
                ColumnsToDateTime = new List<string>(),
                ColumnsToNumber = new List<string>()
            };
        }

        /// <summary>Convert the first column in single worksheet inside an Excel file into a list of strings.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert the first column in single worksheet inside an Excel file into a list of strings.")]
        public List<string> ToListDataString([Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToListDataString(filePath,"",this.DefaultOptions);
        }
        /// <summary>Last error message if one occurred during processing.</summary>
        [Description("Last error message if one occurred during processing.")]
        public string ErrorMessage = "";
        /// <summary>Convert the first column in single worksheet inside an Excel file into a list of strings.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert the first column in single worksheet inside an Excel file into a list of strings.")]
        public List<string> ToListDataString([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Name of the Excel worksheet. If blank will use the first.")]string sheetName)
        {
            return this.ToListDataString(filePath,sheetName,this.DefaultOptions);
        }
        /// <summary>Convert the first column in single worksheet inside an Excel file into a list of strings.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert the first column in single worksheet inside an Excel file into a list of strings.")]
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
        [Description("Convert the first column in single worksheet inside an Excel file into a list of integers.")]
        public List<int> ToListDataInt32([Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToListDataInt32(filePath,"",this.DefaultOptions);
        }
        /// <summary>Convert the first column in single worksheet inside an Excel file into a list of integers.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        [Description("Convert the first column in single worksheet inside an Excel file into a list of integers.")]
        public List<int> ToListDataInt32([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Name of the Excel worksheet. If blank will use the first.")]string sheetName)
        {
            return this.ToListDataInt32(filePath,sheetName,this.DefaultOptions);
        }
        /// <summary>Convert the first column in single worksheet inside an Excel file into a list of integers.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert the first column in single worksheet inside an Excel file into a list of integers.")]
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
            int sheetNumber = this.GetTableIndex(dataSet,sheetName);
            if(sheetNumber < 0)
                sheetNumber = 0;
            listData.AddRange(this.ToListDataInt32(dataSet.Tables[sheetNumber], handlerOptions));
            return listData;
        }

        /// <summary>Convert a single worksheet inside an Excel file into a list of type 'T'.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert a single worksheet inside an Excel file into a list of type 'T'.")]
        public List<T> ToDataTable<T>([Description("Full path and name of the Excel XLSX document.")]string filePath) where T : new()
        {
            return this.ToDataTable<T>(filePath,"",this.DefaultOptions);
        }
        /// <summary>Convert a single worksheet inside an Excel file into a list of type 'T'.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        [Description("Convert a single worksheet inside an Excel file into a list of type 'T'.")]
        public List<T> ToDataTable<T>([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Name of the Excel worksheet. If blank will use the first.")]string sheetName) where T : new()
        {
            return this.ToDataTable<T>(filePath,sheetName,this.DefaultOptions);
        }
        /// <summary>Convert a single worksheet inside an Excel file into a list of type 'T'.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="sheetName">Name of the Excel worksheet. If blank will use the first.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a single worksheet inside an Excel file into a list of type 'T'.")]
        public List<T> ToDataTable<T>([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Name of the Excel worksheet. If blank will use the first.")]string sheetName, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new()
        {
            ErrorMessage = "";
            byte[] byteArray = this.ToExcelBinary(filePath, handlerOptions);
            List<T> listData = new List<T>();
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
            listData.AddRange(this.ToListData<T>(dataSet.Tables[sheetNumber], handlerOptions));
            return listData;
        }

        [Description("Find the index of a DataTable contained in a DataSet from its name.")]
        private int GetTableIndex([Description("DataSet to search.")]DataSet dataSet, [Description("Name of the DataTable.")]string tableName)
        {
            if(tableName=="")
                return -1;
            int sheetNumber = -1;
            for(int i = 0; i < dataSet.Tables.Count; i++)
            {
                if(dataSet.Tables[i]?.TableName == tableName)
                {
                    sheetNumber = i;
                    break;
                }
            }
            return sheetNumber;
        }

        /// <summary>Convert a binary data Excel file into a list of type 'T'.</summary>
        /// <param name="byteArray">Excel binary data to convert.</param>
        [Description("Convert a binary data Excel file into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("Excel binary data.")]byte[] byteArray) where T : new()
        {
            return ToListData<T>(byteArray,this.DefaultOptions);
        }
        /// <summary>Convert a binary data Excel file into a list of type 'T'.</summary>
        /// <param name="byteArray">Excel binary data to convert.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert a binary data Excel file into a list of type 'T'.")]
        public List<T> ToListData<T>([Description("Excel binary data.")]byte[] byteArray, [Description("Optional settings.")]ExcelToDataOptions handlerOptions) where T : new()
        {
            ErrorMessage = "";
            List<T> listData = new List<T>();
            DataSet dataSet = this.ToDataSet(byteArray, handlerOptions);
            if(ErrorMessage!="")
                return listData;
            if(dataSet.Tables.Count==0)
            {
                ErrorMessage = $"No sheets were found in DataSet";
                return listData;
            }
            listData.AddRange(this.ToListData<T>(dataSet.Tables[0], handlerOptions));
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
            List<T> listData = new List<T>();
            if(dataSet.Tables.Count==0)
            {
                ErrorMessage = $"No sheets were found in DataSet";
                return listData;
            }
            int sheetNumber = this.GetTableIndex(dataSet,tableName);
            if(sheetNumber < 0)
                sheetNumber = 0;
            listData.AddRange(this.ToListData<T>(dataSet.Tables[0], handlerOptions));
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
            WorkbookStylesPart stylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
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

        /// <summary>Convert an Excel file into a DataSet.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert an Excel file into a DataSet.")]
        public DataSet ToDataSet([Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToDataSet(filePath,this.DefaultOptions);
        }
        /// <summary>Convert an Excel file into a DataSet.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        /// <param name="handlerOptions">Settings to modify this action.</param>
        [Description("Convert an Excel file into a DataSet.")]
        public DataSet ToDataSet([Description("Full path and name of the Excel XLSX document.")]string filePath, [Description("Optional settings.")]ExcelToDataOptions handlerOptions)
        {
            byte[] byteArray = this.ToExcelBinary(filePath,handlerOptions);
            if(this.ErrorMessage!="")
                return new DataSet();
            DataSet dataSet = this.ToDataSet(byteArray,handlerOptions);
            if(dataSet.Tables.Count==0 && this.ErrorMessage == "")
                this.ErrorMessage = $"No data was found in '{filePath}'";
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
                        foreach (Sheet sheet in sheets)
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
                            if(useHeadings)
                            {
                                Row headerRow = sheetData.Descendants<Row>().First();
                                if(headerRow.HasChildren==true)
                                {
                                    foreach (Cell cell in headerRow.Descendants<Cell>())
                                    {
                                        string columnName = GetCellValue(workbookPart, cell);
                                        if(!this.IsValidWorksheetName(columnName))
                                        {
                                            if(columnIndex==0)
                                            {
                                                columnName = defaultColumnName;
                                            }
                                            else
                                            {
                                                columnName = System.Guid.NewGuid().ToString().Replace("-","").ToLower().Substring(0,31);
                                            }
                                        }
                                        dataTable.Columns.Add(new DataColumn(columnName, typeof(string)));
                                        columnIndex++;
                                        columnTotal++;
                                    }
                                }
                                else
                                {
                                    dataTable.Columns.Add(new DataColumn(defaultColumnName, typeof(string)));
                                    columnIndex++;
                                }
                            }
                            else
                            {
                                skipFirst = 0;
                                Row firstRow = sheetData.Descendants<Row>().First();
                                for(int i = 0; i<firstRow.Descendants().Count(); i++)
                                {
                                    string columnName = defaultColumnName.TrimEnd('1') + i.ToString().PadLeft(3,'0');
                                    dataTable.Columns.Add(new DataColumn(columnName, typeof(string)));
                                    columnIndex++;
                                }
                            }

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

//                                foreach (Cell cell in row.Elements<Cell>())
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
                                        dataRow[columnIndex] = "";
                                    }
                                    else
                                    {
                                        string cellValue = (GetCellValue(workbookPart, cell) ?? "").ToString();
                                        dataRow[columnIndex] = cellValue;
                                    }
                                    columnIndex++;
                                }
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
            int index = 0;
            // string cellReference = cell.CellReference.ToString().ToUpper();
            foreach (char ch in cellReference)
            {
                if (Char.IsLetter(ch))
                {
                    int value = (int)ch - (int)'A';
                    index = (index == 0) ? value : ((index + 1) * 26) + value;
                }
                else
                {
                    return index;
                }
            }
            return index;
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

        /// <summary>Convert the first worksheet inside an Excel file into a DataTable.</summary>
        /// <param name="filePath">Full path and name of the Excel XLSX document.</param>
        [Description("Convert the first worksheet inside an Excel file into a DataTable.")]
        public DataTable ToDataTable([Description("Full path and name of the Excel XLSX document.")]string filePath)
        {
            return this.ToDataTable(filePath,"",this.DefaultOptions);
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
            if(!dataSet.Tables.Contains(sheetName))
            {
                if(this.ErrorMessage == "")
                    this.ErrorMessage = $"Excel file '{filePath}' does not contain workseet '{sheetName}'";
                return new DataTable();
            }
            return dataSet.Tables[sheetName];
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
                    string itemvalue = Fields[i-Props.Length]?.GetValue(item)?.ToString() ?? "";
                    if(itemvalue == "")
                    {
                        values[i] = "";
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
            List<T> listData = new List<T>();
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
                        destValue = this.ParseExcel(sourceValue, destType, handlerOptions);
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
    }
}