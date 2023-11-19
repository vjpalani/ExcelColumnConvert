using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelConverter.Core;
using ExcelConverter.Adapters;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ExcelConverter.Model;

namespace ApiAdaptor.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelConverterController : ControllerBase
    {
        private readonly ILogger<ExcelConverterController> _logger;
        private readonly IExcelConverterAdapter _excelConverter;

        public ExcelConverterController(ILogger<ExcelConverterController> logger, IExcelConverterAdapter excelConverter)
        {
            _logger = logger;
            _excelConverter = excelConverter;
        }

        // Converts a column in an Excel file to its corresponding letter representation.
        //
        // Parameters:
        //   columnDetail:
        //     The details of the column including the columnName and ColumnLetter   .
        //
        // Returns:
        //   An IActionResult containing the result of the conversion as an ApiResponse.
        //   If the conversion is successful, the result will be returned as Ok.
        //   If an error occurs during the conversion process, a BadRequest response
        //   with the error message "Invalid column number." will be returned.

        [HttpPost("columnConverter")]
        public IActionResult ColumnConverter([FromBody] ColumnDetail columnDetail)
        {
            try
            {
                ApiResponse result = _excelConverter.ConvertExcelColumn(columnDetail);
                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error converting column name to letter.");
                return BadRequest("Invalid column number.");
            }
        }
    }
}