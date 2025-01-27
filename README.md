# ExcelToData

## Description
A general-purpose XLSX-to-.NET data conversion tool.
Sample code to manipulate List<t>, byte[], and DataTable objects is below.

## Project Details
This tool has been built to solve ad hoc problems that arise in life involving Excel documents.  Microsoft Excel is the 'hammer and tongs of the Information Age'.  Robust, repeatable, open-source, and memory-safe manipulation of XLSX files is still a handy trick even in a post-LLM world.

## Usage
A compiled version of this sits on Nuget:
https://www.nuget.org/packages/Seamlex.Utilities.ExcelToData

## Installation
This is a class library that can be added to a project with:
    dotnet add package Seamlex.Utilities.ExcelToData

## Sample code
    #pragma warning disable CS1998
    internal class Program
    {
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

## Updating GitHub + NuGet
To create a new version:
1 Update the CHANGELOG.md with the new version
2 Update the PackageVersion field in ExcelToData.csproj
3 Run this in the CLI to push to GitHub:
    dotnet build
    git add .
    git commit -m "1.x.y Fixed XYZ issue"
    git push
4 Run this in the CLI to push to Nuget:
    dotnet pack
    dotnet nuget push {sourcepath}\ExcelToData\bin\Debug\Seamlex.Utilities.ExcelToData.{version in format "1.x.y"}.nupkg --api-key {APIkey} --source https://api.nuget.org/v3/index.json

