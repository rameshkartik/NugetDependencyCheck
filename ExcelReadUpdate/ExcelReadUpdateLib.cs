using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelReadUpdate.Model;
using Newtonsoft.Json;
using NuGet.Versioning;
using NugetRead.Library;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReadUpdate.Library
{
    public class ExcelReadUpdate
    {
        ExcelPackageList excelPackageList = new ExcelPackageList();
        List<ExcelPackageList> listPackages = new List<ExcelPackageList>();
        StringBuilder sbPackagesJson = new StringBuilder();
        private string InputFileName = Assembly.GetExecutingAssembly().Location.Substring(0, Assembly.GetExecutingAssembly().Location.LastIndexOf("\\")) + "\\PackageList.xlsx";
        private readonly string JsonFileName = Assembly.GetExecutingAssembly().Location.Substring(0, Assembly.GetExecutingAssembly().Location.LastIndexOf("\\")) + "\\PackageListJson.txt";
        private readonly string OutputFileName = Assembly.GetExecutingAssembly().Location.Substring(0, Assembly.GetExecutingAssembly().Location.LastIndexOf("\\")) + "\\PackageListOutput.xlsx";
        private readonly string sPackageNameColumnReference = "A";
        NugetReadLib NugetReadLib = new NugetReadLib();
        public ExcelReadUpdate() { }


        public void ReadExcelToPrepareJSONAndUpdateExcel(string sInputFile)
        {
            InputFileName = sInputFile;
            ReadPackageDetails();

        }

        public void ReadPackageDetails()
        {
            using (FileStream fs = new FileStream(InputFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    Worksheet sheet = worksheetPart.Worksheet;

                    var cells = sheet.Descendants<Cell>();
                    var rows = sheet.Descendants<Row>();

                    //Console.WriteLine("Row count = {0}", rows.LongCount());
                    //Console.WriteLine("Cell count = {0}", cells.LongCount());

                   var result =  IterateRowsAndCells(rows, sst);

                    Console.WriteLine(result.ToString());
                }
            }
        }

        private async Task<string> IterateRowsAndCells(IEnumerable<Row> rows, SharedStringTable sst)
        {
            ExcelPackageList packageList;

            foreach (Row row in rows)
            {
                foreach (Cell c in row.Elements<Cell>())
                {
                    if (c.CellReference.ToString().StartsWith(sPackageNameColumnReference) && !c.CellReference.ToString().EndsWith("1"))
                    {
                        packageList = new ExcelPackageList();

                        if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(c.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            packageList.Name = str;
                            var version = await NugetReadLib.FetchLatestNugetVersion(str);
                            var targetFrameworks = await NugetReadLib.GetNugetTargetFrameworks(str, version);
                            var fw80 = targetFrameworks?.ToList().Where(c1 => c1.TargetFramework.ToString().Contains("8.0"));
                            var fw20 = targetFrameworks?.ToList().Where(c1 => c1.TargetFramework.ToString().Contains("2.0"));
                            var fw21 = targetFrameworks?.ToList().Where(c1 => c1.TargetFramework.ToString().Contains("2.1"));
                            packageList.Net80SupportAvailability = fw80?.ToList().Count() > 0 ? "Available" : "Not Available" ;
                            packageList.Net20SupportAvailability = fw20?.ToList().Count() > 0 ? "Available" : "Not Available"; 
                            packageList.Net21SupportAvailability = fw21?.ToList().Count() > 0 ? "Available" : "Not Available";
                            listPackages.Add(packageList);
                        }
                        else if (c.CellValue != null)
                        {
                            packageList.Name = c.CellValue.Text;
                            packageList.Net80SupportAvailability = "Yes";
                            packageList.Net20SupportAvailability = "No";
                            packageList.Net21SupportAvailability = "No";
                            listPackages.Add(packageList);
                        }
                    }
                }

                
            }
            string jsonData = JsonConvert.SerializeObject(listPackages.ToArray());
            sbPackagesJson.Append(jsonData);
            System.IO.File.WriteAllText(JsonFileName, sbPackagesJson.ToString());
            LoadJSONContent(JsonFileName);
            return await Task.FromResult(Task.CompletedTask.ToString());
        }

        private void LoadJSONContent(string jsonFileName)
        {
            using (StreamReader r = new StreamReader(jsonFileName))
            {
                string json = r.ReadToEnd();
                List<ExcelPackageList> items = JsonConvert.DeserializeObject<List<ExcelPackageList>>(json);
                IterateJSONAndUpdateExcel(items);
            }
        }


        private void IterateJSONAndUpdateExcel(List<ExcelPackageList> items)
        {
            uint iCount = 1;
            foreach (ExcelPackageList item in items)
            {
                iCount++;
                UpdateCell(OutputFileName, item.Net80SupportAvailability, iCount, "B");
                UpdateCell(OutputFileName, item.Net20SupportAvailability, iCount, "C");
                UpdateCell(OutputFileName, item.Net21SupportAvailability, iCount, "D");

            }
        }

        public static void UpdateCell(string docName, string text,
           uint rowIndex, string columnName)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet =
                     SpreadsheetDocument.Open(docName, true))
            {
                WorksheetPart worksheetPart =
                      GetWorksheetPartByName(spreadSheet, "Sheet1");

                if (worksheetPart != null)
                {
                    Cell cell = GetCell(worksheetPart.Worksheet,
                                             columnName, rowIndex);

                    cell.CellValue = new CellValue(text);
                    cell.DataType =
                        new EnumValue<CellValues>(CellValues.String);

                    // Save the worksheet.
                    worksheetPart.Worksheet.Save();
                }
            }

        }

        private static WorksheetPart
             GetWorksheetPartByName(SpreadsheetDocument document,
             string sheetName)
        {
            IEnumerable<Sheet> sheets =
               document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.

                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)
                 document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;

        }

        // Given a worksheet, a column name, and a row index, 
        // gets the cell at the specified column and 
        private static Cell GetCell(Worksheet worksheet,
                  string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>().Where(c => string.Compare
                   (c.CellReference.Value, columnName +
                   rowIndex, true) == 0).First();
        }


        // Given a worksheet and a row index, return the row.
        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }


    }
}
