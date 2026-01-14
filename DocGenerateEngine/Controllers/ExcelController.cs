using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using System.IO;

public class ExcelController : Controller
{
    // GET: Upload page
    [HttpGet]
    public IActionResult Upload()
    {
        return View(); // This still returns the Upload.cshtml page
    }

    // POST: File upload
    [HttpPost]
    public IActionResult Upload(IFormFile excelFile)
    {
        if (excelFile == null || excelFile.Length == 0)
            return BadRequest("No file selected"); // JS will handle this

        var uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/uploads");
        if (!Directory.Exists(uploadPath))
            Directory.CreateDirectory(uploadPath);

        var filePath = Path.Combine(uploadPath, excelFile.FileName);
        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            excelFile.CopyTo(stream);
        }

        // Return simple success message for JS
        return Ok("File uploaded successfully");
    }
}
