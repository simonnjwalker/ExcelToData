# ExcelToData
Convert to-and-from simple Excel files and .NET data objects

# Usage
Sample code to manipulate List<t>, byte[], and DataTable objects is below.

# Version History


-----------
1.0.4 Added CSV input/output
-----------

8 - Added CSV readinhg/writing
https://github.com/simonnjwalker/ExcelToData/issues/8



-----------
1.0.3 Fixed issues where OpenXML skips blank cells and infering tablename from type
-----------

Fixed to correctly skip blank cells in headers and in data
https://github.com/simonnjwalker/ExcelToData/issues/6

7 - fixed to should infer the worksheet name <T> when List<T> is the required output
https://github.com/simonnjwalker/ExcelToData/issues/7



-----------
1.0.2 Fixed columns and allowed indexes for sheets
-----------

4 - fixed to check header row for merged, missing, and duplicated field names
https://github.com/simonnjwalker/ExcelToData/issues/4

5 - added sheet indexes to ToDataTable()
https://github.com/simonnjwalker/ExcelToData/issues/5



-----------
1.0.1 Fixed bugs:
-----------

1 - name of method public List<T> ToDataTable<T> should be ToListData<T>
- fixed 

2 - GetDefaults() returns the internal settings object and not a clone
- removed GetDefaults()
- added GetOptionsClone()

3 - Error where List<T> has non-CLR types
- fixed
- will convert to-and-from Json and put this into the xlsx column
- this is the default behaviour, change it with:
xlsx.GetOptions().ComplexToJson = false;



-----------
1.0.0 Initial release
-----------





# Sample code

// this is a dotnet 7 console application
#pragma warning disable CS1998
internal class Program
{
    // the Main method will complain about lacking awaits
    private static async Task Main(string[] args)
    {
        // set sample file names (MS Windows)
        string output = @"c:\temp\test1.xlsx";
        string output2 = @"c:\temp\test2.xlsx";

        // instantiate this class
        var xlsx = new Seamlex.Utilities.ExcelToData();

        // create some demo data in a List<T>
        var clients = new List<MyClass>(){
            new MyClass() { IsValid = false, DollarAmount = 123.45, StartDate = DateTime.Now.AddDays(7), EndDate = DateTime.Now.AddMonths(3), FirstName = "Asha", LastName = "Albatross", MyProp = new ClassProp(){IsValid=true, DollarAmount=2.50}  },
            new MyClass() { IsValid = true, DollarAmount = 400.00, StartDate = DateTime.Now.AddDays(14), EndDate = DateTime.Now.AddMonths(4),FirstName = "Bianca", LastName = "Best", MyProp = new ClassProp(){IsValid=false, DollarAmount=0.50}  },
            new MyClass() { IsValid = false, DollarAmount = 100.00, StartDate = DateTime.Now.AddDays(21), EndDate = DateTime.Now.AddMonths(5), FirstName = "Carl", LastName = "Cranston", MyProp = new ClassProp(){IsValid=true, DollarAmount=1.50} }
        };

        // this changes default behaviour of the client which outputs as text to all fields 
        //xlsx.GetOptions().ColumnsToDateTime.Add("StartDate");
        //xlsx.GetOptions().ColumnsToNumber.Add("DollarAmount");

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
            $"{item.FirstName} {item.LastName} {item.StartDate} {item.EndDate} {item.DollarAmount} {item.MyProp.DollarAmount} ");

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

    public ClassProp MyProp = new ClassProp();
    
}    

public class ClassProp {
    public bool IsValid = false;
    public double DollarAmount { get; set; }
    
}    

    
