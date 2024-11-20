using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Collections.Concurrent;

namespace ExcelValidationApp.Pages
{
    public class UploadExcelModel : PageModel
    {
        [BindProperty]
        public IFormFile? File { get; set; }

        public string? Message { get; set; }
        public bool IsError { get; set; }

        public List<Dictionary<string, string>> ProcessedData { get; set; } = new();

        private readonly HarbColumnDictionary _harbColumnDictionary;

        // Constructor injection for the singleton HarbColumnDictionary
        public UploadExcelModel(HarbColumnDictionary harbColumnDictionary)
        {
            _harbColumnDictionary = harbColumnDictionary;
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (File == null || File.Length == 0)
            {
                Message = "No file uploaded!";
                IsError = true;
                return Page();
            }

            try
            {
                using var stream = File.OpenReadStream();
                using var workbook = new XLWorkbook(stream);
                var worksheet = workbook.Worksheet(1); // Read the first worksheet

                // Skip the first two rows to find the header
                int headerRow = 4;
                var headerCells = worksheet.Row(headerRow).CellsUsed().ToList();
                if(headerCells.Count == 0)
                {
                    Message = "Column validation failed. no title found";
                    IsError = true;
                    return Page();
                }

                // Get the column names from the header row
                var columnNames = headerCells.Select(c => c.GetString()).ToList();

                // Validate column indexes and display names
                if (!ValidateColumns(columnNames))
                {
                    Message = "Column validation failed. Check the uploaded file.";
                    IsError = true;
                    return Page();
                }

                // Get the index of the "Main Branch" column using HarbColumn.MainBranchIndex
                int mainBranchColumnIndex = _harbColumnDictionary.GetIndexByName(nameof(HarbRow.MainBranch));
                if (mainBranchColumnIndex == -1)
                {
                    Message = "Column 'Main Branch' not found!";
                    IsError = true;
                    return Page();
                }

                var harbRows = new List<HarbRow>();

                // Process rows, skipping section headers and empty rows
                foreach (var row in worksheet.RowsUsed().Skip(headerRow))
                {
                    var mainBranchValue = row.Cell(mainBranchColumnIndex + 1).GetString();

                    // Skip section headers (rows where "Main Branch" starts with "ъзен")
                    if (mainBranchValue.StartsWith("ъзен"))
                        continue;



                    var harbRow = new HarbRow
                    {
                        MainBranch = row.Cell(_harbColumnDictionary.GetIndexByName(nameof(HarbRow.MainBranch)) + 1).GetString(),
                        SubBranch = row.Cell(_harbColumnDictionary.GetIndexByName(nameof(HarbRow.SubBranch)) + 1).GetString(),
                        ProductType = row.Cell(_harbColumnDictionary.GetIndexByName(nameof(HarbRow.ProductType)) + 1).GetString(),
                        Company = row.Cell(_harbColumnDictionary.GetIndexByName(nameof(HarbRow.Company)) + 1).GetString(),
                        InsurancePeriod = row.Cell(_harbColumnDictionary.GetIndexByName(nameof(HarbRow.InsurancePeriod)) + 1).GetString(),
                        Premium = row.Cell(_harbColumnDictionary.GetIndexByName(nameof(HarbRow.Premium)) + 1).GetString(),
                        PremiumType = row.Cell(_harbColumnDictionary.GetIndexByName(nameof(HarbRow.PremiumType)) + 1).GetString(),
                        PolicyNumber = row.Cell(_harbColumnDictionary.GetIndexByName(nameof(HarbRow.PolicyNumber)) + 1).GetString(),
                        PlanClassification = row.Cell(_harbColumnDictionary.GetIndexByName(nameof(HarbRow.PlanClassification)) + 1).GetString()
                    };

                    harbRows.Add(harbRow);

                    // Check for duplicate SubBranch values
                    var duplicates = harbRows.GroupBy(r => r.SubBranch).Where(g => g.Count() > 1).ToList();
                    if (duplicates.Any())
                    {
                        Message = "duplicates rows";
                        IsError = true;
                        return Page();
                    }


                    // Build a data dictionary for this row
                    var rowData = new Dictionary<string, string>();
                    foreach (var column in _harbColumnDictionary.Columns)
                    {
                        var cellValue = row.Cell(column.Key+ 1).GetString();
                        rowData[column.Value.DisplayName] = cellValue;
                    }

                    if (rowData.Count > 0)
                    {
                        ProcessedData.Add(rowData);
                    }
                }

                Message = "File uploaded and processed successfully!";
                IsError = false;
                return Page();
            }
            catch (Exception ex)
            {
                Message = $"Error processing file: {ex.Message}";
                IsError = true;
                return Page();
            }
        }

        // Validate column indexes and display names
        private bool ValidateColumns(List<string> columnNames)
        {
            foreach (var column in _harbColumnDictionary.Columns)
            {
                if (column.Key >= columnNames.Count || columnNames[column.Key] != column.Value.DisplayName)
                {
                    return false;
                }
            }
            return true;
        }
    }
}





