using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"D:\New folder\App.xlsx");

            var labor = GetSetupData();
            await SaveExcelFile(labor, file);
            List<LaborproModel> laborprosFromexcel = await LoadExcelFile(file);

            Console.WriteLine("Id   FeatureFileName    ScenarioName      SmokeTest    RegressionTest");
            foreach(var l in laborprosFromexcel)
            {
                
                Console.WriteLine($"{l.Id} {l.FeatureFileName} {l.ScenarioName} {l.SmokeTest} {l.RegressionTest}");
            }
       
        }

        private static async Task<List<LaborproModel>> LoadExcelFile(FileInfo file)
        {
            List<LaborproModel> output = new();
            using var package = new ExcelPackage(file);
            await package.LoadAsync(file);
            var ws = package.Workbook.Worksheets[0];
            int row = 3;
            int col = 1;
            while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
            {
                LaborproModel l = new();
                l.Id = int.Parse(ws.Cells[row,col].Value.ToString());
                l.FeatureFileName = ws.Cells[row,col+1].Value.ToString();
                l.ScenarioName = ws.Cells[row, col+2].Value.ToString();
                l.SmokeTest = ws.Cells[row, col+3].Value.ToString();
                l.RegressionTest = ws.Cells[row,col+4].Value.ToString();
                output.Add(l);
                row += 1;
            }
            return output;
        }

        private static async Task SaveExcelFile(List<LaborproModel> labor, FileInfo file)
        {
           DeleteIfExist(file);
            using var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("TestReport");
            var range = ws.Cells["A2"].LoadFromCollection(labor, true);
            range.AutoFitColumns();
            ws.Cells["A1"].Value = "Test Report";
            ws.Cells["A1:C1"].Merge = true;
            ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(1).Style.Font.Size = 24;
            ws.Row(1).Style.Font.Color.SetColor(Color.Blue);
            
            ws.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(2).Style.Font.Bold = true;
            ws.Column(3).Width = 20;
            

            await package.SaveAsync();


        }

        private static void DeleteIfExist(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private static List<LaborproModel> GetSetupData()
        {
            List<LaborproModel> output = new()
            {
                new() { Id =1, FeatureFileName = "Demo", ScenarioName = "Demo", SmokeTest = "yes", RegressionTest="no"},
                new() { Id = 1, FeatureFileName = "Demo", ScenarioName = "Demo", SmokeTest = "yes", RegressionTest = "no" },
                new() { Id = 1, FeatureFileName = "Demo", ScenarioName = "Demo", SmokeTest = "yes", RegressionTest = "no" },
                new() { Id = 1, FeatureFileName = "Demo", ScenarioName = "Demo", SmokeTest = "yes", RegressionTest = "no" },
                new() { Id = 1, FeatureFileName = "Demo", ScenarioName = "Demo", SmokeTest = "yes", RegressionTest = "no" },
                new() { Id = 1, FeatureFileName = "Demo", ScenarioName = "Demo", SmokeTest = "yes", RegressionTest = "no" },

            };
            return output;
        }

    }
}
