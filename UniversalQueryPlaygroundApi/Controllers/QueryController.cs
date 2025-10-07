using Microsoft.AspNetCore.Mvc;

namespace UniversalQueryPlaygroundApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class WorkbookController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly string _uploadFolder;
        private readonly string _activeWorkbookPath;

        public WorkbookController(IConfiguration config)
        {
            _config = config;
            _uploadFolder = @"C:\Users\kayod\Desktop\DataTmp";
            _activeWorkbookPath = _config["DataSources:ExcelFile"] 
                ?? throw new Exception("DataSources:ExcelFile not configured in appsettings.json.");

            if (!Directory.Exists(_uploadFolder))
                Directory.CreateDirectory(_uploadFolder);
        }

        /// <summary>
        /// Returns the current active workbook file.
        /// </summary>
        [HttpGet("latest")]
        public IActionResult GetLatestWorkbook()
        {
            if (!System.IO.File.Exists(_activeWorkbookPath))
                return NotFound("No active workbook found.");

            var bytes = System.IO.File.ReadAllBytes(_activeWorkbookPath);
            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                Path.GetFileName(_activeWorkbookPath));
        }

        /// <summary>
        /// Uploads a new Excel workbook to the server, stores it in the upload folder,
        /// and sets it as the active workbook by copying to the configured path.
        /// </summary>
        [HttpPost("upload")]
        [RequestSizeLimit(50_000_000)] // ~50 MB
        public async Task<IActionResult> UploadWorkbook(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xls")
                return BadRequest("Only Excel files (.xlsx, .xls) are supported.");

            // Save original upload with timestamp for reference
            var timestamp = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss");
            var uploadFileName = $"{Path.GetFileNameWithoutExtension(file.FileName)}_{timestamp}{ext}";
            var uploadPath = Path.Combine(_uploadFolder, uploadFileName);

            using (var stream = new FileStream(uploadPath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            // Copy to active workbook location
            System.IO.File.Copy(uploadPath, _activeWorkbookPath, overwrite: true);

            return Ok(new
            {
                message = "Workbook uploaded successfully.",
                originalName = file.FileName,
                uploadPath,
                activeWorkbook = _activeWorkbookPath
            });
        }
    }
}
