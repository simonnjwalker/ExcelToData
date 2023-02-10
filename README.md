# ExcelToData
Convert to-and-from simple Excel files and .NET data objects

# Usage
Sample code to manipulate List<t>, byte[], and DataTable objects is below.

# Sample code

// this is a dotnet 7 console application
internal class Program {
private static async Task Main(string[] args) {
    
// set sample file names (MS Windows)
string output = @"c:\temp\test1.xlsx";
string output2 = @"c:\temp\test2.xlsx";

// instantiate this class
var xlsx = new Seamlex.Utilities.ExcelToData();

// create some demo data in a List<T>
var clients = new List<MyClass>(){
    new MyClass() { IsValid = false, DollarAmount = 123.45, StartDate = DateTime.Now.AddDays(7), EndDate = DateTime.Now.AddMonths(3), FirstName = "Asha", LastName = "Albatross" },
    new MyClass() { IsValid = true, DollarAmount = 400.00, StartDate = DateTime.Now.AddDays(14), EndDate = DateTime.Now.AddMonths(4),FirstName = "Bianca", LastName = "Best" },
    new MyClass() { IsValid = false, DollarAmount = 100.00, StartDate = DateTime.Now.AddDays(21), EndDate = DateTime.Now.AddMonths(5), FirstName = "Carl", LastName = "Cranston" }
};

// this changes default behaviour of the client which outputs as text to all
// fields comment these four lines and the console output at the end will be identical
xlsx.GetDefaults().ColumnsToDateTime.Add("StartDate");
xlsx.GetDefaults().ColumnsToNumber.Add("DollarAmount");

// this creates an Excel file from a .NET List<T>
xlsx.ToExcelFile(clients, output);

// this displays the last error (if any)
if (xlsx.ErrorMessage != "")
  Console.WriteLine(xlsx.ErrorMessage);

// this loads an Excel file into a System.Data.DataTable object
var dt = xlsx.ToDataTable(output);

// this saves it to a second file
xlsx.ToExcelFile(dt, output2);

// this reloads that second file into an in-memory byte array
byte[] file = xlsx.ToExcelBinary(output2);

// this takes the in-memory object and converts it to a List<T> then outputs
// to the console
var list2 = xlsx.ToListData<MyClass>(file);
foreach (var item in list2)
  Console.WriteLine(
      $"{item.FirstName} {item.LastName} {item.StartDate} {item.EndDate}");

return;
}
}

// demo class with a combination of fields and properties
public class MyClass {
public bool IsValid = false;
public double DollarAmount { get; set; }
public DateTime StartDate = DateTime.Now;
public DateTime? EndDate { get; set; }
public string FirstName = "";
public string LastName = "";
}

