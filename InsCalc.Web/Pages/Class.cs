//using ClosedXML.Excel;
//using Microsoft.AspNetCore.Mvc;

//[ApiController]
//[Route("[controller]")]
//public class FileUploadController : ControllerBase
//{
//    private readonly HarbColumnDictionary _harbColumnDictionary;

//    public FileUploadController(HarbColumnDictionary harbColumnDictionary)
//    {
//        _harbColumnDictionary = harbColumnDictionary;
//    }

//    [HttpPost("upload")]
//    public IActionResult Upload(IFormFile file)
//    {
//        if (file == null || Path.GetExtension(file.FileName).ToLower() != ".xlsx")
//            return BadRequest("Invalid file. Please upload an Excel (.xlsx) file.");

//        try
//        {
//            using var stream = new MemoryStream();
//            file.CopyTo(stream);

//            using var workbook = new XLWorkbook(stream);
//            var worksheet = workbook.Worksheets.First();

//            // Skip the first two rows to find the header
//            var headerRow = worksheet.Row(3);
//            var columnNames = headerRow.CellsUsed().Select(c => c.GetString()).ToList();

//            // Populate HarbColumnDictionary
//            for (var i = 0; i < columnNames.Count; i++)
//            {
//                _harbColumnDictionary.Add(columnNames[i], i);
//            }

//            // Validate columns
//            var expectedColumns = new List<(string name, string displayName)>
//            {
//                (nameof(HarbRow.MainBranch), "ענף ראשי"),
//                (nameof(HarbRow.SubBranch), "ענף (משני)"),
//                (nameof(HarbRow.ProductType), "סוג מוצר"),
//                (nameof(HarbRow.Company), "חברה"),
//                (nameof(HarbRow.InsurancePeriod), "תקופת ביטוח"),
//                (nameof(HarbRow.Premium), "פרמיה בש\"ח"),
//                (nameof(HarbRow.PremiumType), "סוג פרמיה"),
//                (nameof(HarbRow.PolicyNumber), "מספר פוליסה"),
//                (nameof(HarbRow.PlanClassification), "סיווג תכנית")
//            };

//            if (!_harbColumnDictionary.ValidateColumns(expectedColumns))
//                return BadRequest("Invalid columns in the uploaded file.");

//            // Parse rows into HarbRow objects
//            var harbRows = new List<HarbRow>();
//            foreach (var row in worksheet.RowsUsed().Skip(3))
//            {
//                var mainBranchValue = row.Cell(_harbColumnDictionary.GetIndexByKey(nameof(HarbRow.MainBranch)) + 1).GetString();

//                // Skip section headers
//                if (mainBranchValue.StartsWith("תחום"))
//                    continue;

//                // Create HarbRow object
//                var harbRow = new HarbRow
//                {
//                    MainBranch = row.Cell(_harbColumnDictionary.GetIndexByKey(nameof(HarbRow.MainBranch)) + 1).GetString(),
//                    SubBranch = row.Cell(_harbColumnDictionary.GetIndexByKey(nameof(HarbRow.SubBranch)) + 1).GetString(),
//                    ProductType = row.Cell(_harbColumnDictionary.GetIndexByKey(nameof(HarbRow.ProductType)) + 1).GetString(),
//                    Company = row.Cell(_harbColumnDictionary.GetIndexByKey(nameof(HarbRow.Company)) + 1).GetString(),
//                    InsurancePeriod = row.Cell(_harbColumnDictionary.GetIndexByKey(nameof(HarbRow.InsurancePeriod)) + 1).GetString(),
//                    Premium = row.Cell(_harbColumnDictionary.GetIndexByKey(nameof(HarbRow.Premium)) + 1).GetString(),
//                    PremiumType = row.Cell(_harbColumnDictionary.GetIndexByKey(nameof(HarbRow.PremiumType)) + 1).GetString(),
//                    PolicyNumber = row.Cell(_harbColumnDictionary.GetIndexByKey(nameof(HarbRow.PolicyNumber)) + 1).GetString(),
//                    PlanClassification = row.Cell(_harbColumnDictionary.GetIndexByKey(nameof(HarbRow.PlanClassification)) + 1).GetString()
//                };

//                harbRows.Add(harbRow);
//            }

//            // Check for duplicate SubBranch values
//            var duplicates = harbRows.GroupBy(r => r.SubBranch).Where(g => g.Count() > 1).ToList();
//            if (duplicates.Any())
//            {
//                return Conflict("The uploaded file contains duplicate SubBranch values.");
//            }

//            return Ok(harbRows);
//        }
//        catch (Exception ex)
//        {
//            return StatusCode(500, $"Internal server error: {ex.Message}");
//        }
//    }
//}
