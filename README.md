
internal class Program
{
    private static async Task Main(string[] args)
    {
        string output = @"c:\temp\test1.xlsx";
        string output2 = @"c:\temp\test2.xlsx";
        var xlsx = new Seamlex.Utilities.ExcelToData();
        var clients = new List<MyClass>();
        clients.Add(new MyClass(){
            IsValid = false,
            DollarAmount = 123.45,
            StartDate = DateTime.Now.AddDays(7),
            EndDate = DateTime.Now.AddMonths(3),
            FirstName = "Asha",
            LastName = "Albatross"
        });
        clients.Add(new MyClass(){
            IsValid = true,
            DollarAmount = 400.00,
            StartDate = DateTime.Now.AddDays(14),
            EndDate = DateTime.Now.AddMonths(4),
            FirstName = "Bianca",
            LastName = "Best"
        });
        clients.Add(new MyClass(){
            IsValid = false,
            DollarAmount = 100.00,
            StartDate = DateTime.Now.AddDays(21),
            EndDate = DateTime.Now.AddMonths(5),
            FirstName = "Carl",
            LastName = "Cranston"
        });

        // this changes default behaviour of the client which outputs as text to all fields
        // comment these four lines and the console output at the end will be identical
        var defaults = xlsx.GetDefaults();
        defaults.ColumnsToDateTime.Add("StartDate");
        defaults.ColumnsToNumber.Add("DollarAmount");
        xlsx.SetDefaults(defaults);

        // this creates an Excel file from a .NET List<T>
        xlsx.ToExcelFile(clients,output);

        // this displays the last error (if any)
        if(xlsx.ErrorMessage!="")
          Console.WriteLine(xlsx.ErrorMessage);

        // this loads an Excel file into a System.Data.DataTable object
        var dt = xlsx.ToDataTable(output);

        // this saves it to a second file
        xlsx.ToExcelFile(dt,output2);

        // this reloads that second file into an in-memory byte array
        byte[] file = xlsx.ToExcelBinary(output2);

        // this takes the in-memory object and converts it to a List<T> then outputs to the console
        var list2 = xlsx.ToListData<MyClass>(file);
        foreach(var item in list2)
            Console.WriteLine($"{item.FirstName} {item.LastName} {item.StartDate} {item.EndDate}");

        return;
    }
}

public class MyClass
{
    public bool IsValid = false;
    public double DollarAmount {get;set;}
    public DateTime StartDate = DateTime.Now;
    public DateTime? EndDate {get;set;}
    public string FirstName = "";
    public string LastName = "";
}