using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

using ConsoleTables;
using OfficeOpenXml;

namespace ConsoleApp
{
    class Program
    {
        static async Task Main(string[] args)
        {
            //Epplus NonCommercial License package from the NuGet
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Get FileName as Input
            string fileName;
            Console.Write("Enter File Name :");
            fileName = Console.ReadLine();
            
            // Getting File Path to Save the Dynamically Generated Excel File
            var file = new FileInfo(Path.Combine(Environment.CurrentDirectory.Replace(@"bin\Debug\netcoreapp3.1", @"SaveExcelFile"), fileName + ".xlsx"));

            // Getting data from from a generic list
            var people = GetSetupData();

            // Generating ExcelSheet and saving the data
            await saveExcelFile(people, file);

            Console.WriteLine("");
            Console.WriteLine("-----------------------------------------------------------------");
            Console.WriteLine("                 Retreiving Data From Excel Sheet                 ");
            Console.WriteLine("-----------------------------------------------------------------");

            // Getting the data from the Excel Sheet
            List<PersonModel> peopleFromExcel = await LoadExcelFile(file);

            // Displaying the retrieved data in a tabular format using ConsoleTable package from the NuGet
            var table = new ConsoleTable("Id", "First Name", "Last Name");
            foreach (var p in peopleFromExcel) {
                table.AddRow(p.Id, p.firstName, p.lastName);
            }
            table.Write(Format.Alternative);
            Console.WriteLine();
            Console.ReadLine();

        }

        private static async Task<List<PersonModel>> LoadExcelFile(FileInfo file)
        {
            List<PersonModel> output = new List<PersonModel>();

            // Create a new instance of the ExcelPackage class based on a existing file or creates a new file.
            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            // Read the file from the Work Sheet at index position 0
            var ws = package.Workbook.Worksheets[PositionID: 0];

            // Get data from Column 1 starting from row 2 --- skiping header row
            // And push the retreived data to the generic list
            int row = 2;
            int col = 1;
            while (string.IsNullOrWhiteSpace(ws.Cells[row,col].Value?.ToString()) == false)
            {
                PersonModel p = new PersonModel();
                p.Id = int.Parse(ws.Cells[row, col].Value?.ToString());
                p.firstName = ws.Cells[row, col + 1].Value?.ToString();
                p.lastName = ws.Cells[row, col + 2].Value?.ToString();
                output.Add(p);
                row += 1;
            }

            // Returning the data
            return output;

        }

        private static async Task saveExcelFile(List<PersonModel> people, FileInfo file)
        {
            // For Deleting the File if already exist
            DeleteIfExists(file);

            // Create a new instance of the ExcelPackage class based on a existing file or creates a new file.
            using var package = new ExcelPackage(file);

            // Creating Work Sheet on the above Excel File
            var ws = package.Workbook.Worksheets.Add(Name:"MainReport");

            // Starting Cell Index in the Work Sheet
            var range = ws.Cells[Address: "A1"].LoadFromCollection(people, PrintHeaders: true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists) {
                file.Delete();
            }
        }

        private static List<PersonModel> GetSetupData()
        {
            List<PersonModel> output = new List<PersonModel>()
            {
                new PersonModel(){ Id = 1, firstName="Shanif", lastName="Hussain"},
                new PersonModel(){ Id = 2, firstName="Uma", lastName="Ugale"},
                new PersonModel(){ Id = 3, firstName="Yuaan", lastName="Shanif"}
            };

            return output;

        }

    }
}
