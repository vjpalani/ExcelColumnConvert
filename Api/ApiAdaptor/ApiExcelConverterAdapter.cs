using System.Text.Json;
using ExcelConverter.Core;
using ExcelConverter.Model;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;


namespace ExcelConverter.Adapters
{
    public class ApiExcelConverterAdapter : IExcelConverterAdapter
    {
        private readonly IExcelConverter _excelConverter;

        public ApiExcelConverterAdapter(IExcelConverter excelConverter)
        {
            _excelConverter = excelConverter;
        }

        // <summary>
        /// Converts the given column detail to the corresponding Excel representation.
        /// </summary>
        /// <param name="columnDetail">The column detail to convert.</param>
        /// <returns>The API response containing the converted value.</r
        public ApiResponse ConvertExcelColumn(ColumnDetail columnDetail)
        {
            ApiResponse apiResponse = new ApiResponse();
            try
            {
                if (columnDetail != null && !string.IsNullOrEmpty(columnDetail.columnLetter) && columnDetail.columnLetter.All(char.IsLetter))
                {
                    string Value = (_excelConverter.ConvertColumnLetterToName(columnDetail.columnLetter));
                    apiResponse.statusCode = 200;
                    apiResponse.status = "SUCCESS";
                    apiResponse.result = "The numerical representation corresponding to the letter '" + columnDetail.columnLetter + "' in Excel is  " + Value + ".";
                }
                else if (columnDetail != null && !string.IsNullOrEmpty(columnDetail.columnName) && int.TryParse(columnDetail.columnName, out _) && columnDetail.columnName != "0")
                {
                    string Value = _excelConverter.ConvertColumnNameToLetter(columnDetail.columnName);
                    apiResponse.statusCode = 200;
                    apiResponse.status = "SUCCESS";
                    apiResponse.result = "The letter corresponding to the numerical value " + columnDetail.columnName + " in Excel is  '" + Value + "'.";
                }
                else
                {
                    apiResponse.statusCode = 422;
                    apiResponse.status = "FAILURE";
                    apiResponse.errorMessage = "Invalid column detail provided.";
                }
            }
            catch (JsonException)
            {
                Console.WriteLine($"There is an error processing");
                apiResponse.statusCode = 400;
                apiResponse.status = "Bad Request";
                apiResponse.errorMessage = "Invalid JSON data in the request body.";
            }
            return apiResponse;
        }
    }
}