using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Text.Json;
using System.Text.RegularExpressions;
using Xceed.Words.NET;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
//namespace DocGenerateEngine.Controllers
//{
public class DocumentController : Controller
    {
        private readonly string _uploadPath;

        public DocumentController()
        {
            _uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");

            if (!Directory.Exists(_uploadPath))
                Directory.CreateDirectory(_uploadPath);
        }

        // =====================================
        // GET: Document/Editor
        // Load editor with uploaded Excel files
        // =====================================
        [HttpGet]
        public IActionResult Editor()
        {
            var files = Directory.Exists(_uploadPath)
                ? Directory.GetFiles(_uploadPath)
                           .Where(f => f.EndsWith(".xlsx") || f.EndsWith(".xls"))
                           .Select(Path.GetFileName)
                           .ToList()
                : new List<string>();

            ViewBag.UploadedFiles = files;
            return View();
        }

        // =====================================
        // GET: Document/GetColumns
        // Return column headers for selected Excel file
        // =====================================
        [HttpGet]
        public IActionResult GetColumns(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return BadRequest("File name is required");

            var filePath = Path.Combine(_uploadPath, fileName);

            if (!System.IO.File.Exists(filePath))
                return NotFound("File not found");

            var columns = new List<string>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.FirstOrDefault();
                if (worksheet != null)
                {
                    var headerRow = worksheet.FirstRowUsed();
                    if (headerRow != null)
                    {
                        foreach (var cell in headerRow.CellsUsed())
                        {
                            columns.Add(cell.GetString());
                        }
                    }
                }
            }

            return Json(columns);
        }


    public static string ConvertExcelToJson(string excelFilePath)
    {
        //if (!File.Exists(excelFilePath))
        //    throw new FileNotFoundException("Excel file not found", excelFilePath);

        var rows = new List<Dictionary<string, string>>();

        using (var workbook = new XLWorkbook(excelFilePath))
        {
            var worksheet = workbook.Worksheets.First();
            var headerRow = worksheet.FirstRowUsed();

            var headers = headerRow.CellsUsed()
                                   .Select(c => c.GetString())
                                   .ToList();

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var rowData = new Dictionary<string, string>();

                for (int i = 0; i < headers.Count; i++)
                {
                    rowData[headers[i]] = row.Cell(i + 1).GetValue<string>();
                }

                rows.Add(rowData);
            }
        }

        return JsonSerializer.Serialize(rows, new JsonSerializerOptions
        {
            WriteIndented = true
        });
    }
  


    [HttpPost]
    public IActionResult GenerateDocuments([FromBody] GenerateDocumentsRequest request)
    {
        if (request == null ||
            string.IsNullOrWhiteSpace(request.ExcelFileName) ||
            string.IsNullOrWhiteSpace(request.TemplateHtml))
            return BadRequest("Invalid request");

        var excelPath = Path.Combine(_uploadPath, request.ExcelFileName);

        if (!System.IO.File.Exists(excelPath))
            return NotFound("Excel file not found");

        // 🔹 Convert Excel → JSON → Object
        var json = ConvertExcelToJson(excelPath);

        var rows = JsonSerializer.Deserialize<
            List<Dictionary<string, string>>
        >(json);

        if (rows == null || rows.Count == 0)
            return BadRequest("No data found in Excel");

        // 🔹 Temp folder
        var tempFolder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempFolder);

        // 🔹 Generate DOCX per row
        for (int i = 0; i < rows.Count; i++)
        {
            string docContent = request.TemplateHtml;

            foreach (var column in rows[i])
            {
                string placeholder = $"{{{{{column.Key}}}}}";
                docContent = docContent.Replace(placeholder, column.Value ?? "");
            }

            string plainText = Regex.Replace(docContent, "<.*?>", "\n");

            string docFileName = $"Document_{i + 1}.docx";
            string docPath = Path.Combine(tempFolder, docFileName);

            using (var wordDoc = WordprocessingDocument.Create(
             docPath,
             DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
             {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                foreach (var line in plainText.Split('\n'))
                {
                    var para = new Paragraph(new Run(new Text(line)));
                    body.Append(para);
                }

                mainPart.Document.Save();
            }
        }

        // 🔹 ZIP
        var zipPath = Path.Combine(Path.GetTempPath(), "GeneratedDocs.zip");
        if (System.IO.File.Exists(zipPath))
            System.IO.File.Delete(zipPath);

        System.IO.Compression.ZipFile.CreateFromDirectory(tempFolder, zipPath);

        byte[] fileBytes = System.IO.File.ReadAllBytes(zipPath);
        Directory.Delete(tempFolder, true);

        return File(fileBytes, "application/zip", "GeneratedDocs.zip");
    }



    private int GetColumnIndexByName(ExcelWorksheet ws, string colName)
        {
            for (int col = 1; col <= ws.Dimension.End.Column; col++)
            {
                if (ws.Cells[1, col].Text.Equals(colName, StringComparison.OrdinalIgnoreCase))
                    return col;
            }
            throw new Exception($"Column {colName} not found in Excel");
        }

        // =====================================
        // Request model for GenerateDocuments
        // =====================================
        public class GenerateDocumentsRequest
        {
            public string ExcelFileName { get; set; }
            public string TemplateHtml { get; set; }
        public Dictionary<string, string> ColumnMapping { get; set; }
    }
    }
//}
