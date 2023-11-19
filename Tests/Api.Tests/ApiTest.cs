using NUnit.Framework;
using ExcelConverter.Adapters;
using ExcelConverter.Core;
using ExcelConverter.Model;

namespace ExcelConverter.Tests
{
    [TestFixture]
    public class ApiTest
    {
        private IExcelConverter _excelConverter;

        [SetUp]
        public void Setup()
        {
            _excelConverter = new ExcelConverterClass();
        }

    
        [Test]
        public void ConvertColumnLetterToName_ValidInput_ReturnsCorrectName()
        {
            IExcelConverterAdapter adapter = new ApiExcelConverterAdapter(_excelConverter);
            ColumnDetail columnDetail = new ColumnDetail();
            ApiResponse expected = new ApiResponse();
            expected.statusCode = 200;
            expected.status = "SUCCESS";
            expected.result = "The numerical representation corresponding to the letter 'AB' in Excel is  28.";
            expected.errorMessage = null;
            columnDetail.columnLetter = "AB";
            ApiResponse actual = adapter.ConvertExcelColumn(columnDetail);
            Assert.That(actual, Has.Property(nameof(ApiResponse.statusCode)).EqualTo(expected.statusCode)
                  .And.Property(nameof(ApiResponse.status)).EqualTo(expected.status)
                  .And.Property(nameof(ApiResponse.result)).EqualTo(expected.result)
                  .And.Property(nameof(ApiResponse.errorMessage)).EqualTo(expected.errorMessage));

        }

        [Test]
        public void ConvertColumnNameToLetter_ValidInput_ReturnsCorrectLetter()
        {
            IExcelConverterAdapter adapter = new ApiExcelConverterAdapter(_excelConverter);
            ColumnDetail columnDetail = new ColumnDetail();
            columnDetail.columnName = "28";
            ApiResponse expected = new ApiResponse();
            expected.statusCode = 200;
            expected.status = "SUCCESS";
            expected.result = "The letter corresponding to the numerical value 28 in Excel is  'AB'.";
            expected.errorMessage = null;
            ApiResponse actual = adapter.ConvertExcelColumn(columnDetail);
            Assert.That(actual, Has.Property(nameof(ApiResponse.statusCode)).EqualTo(expected.statusCode)
                  .And.Property(nameof(ApiResponse.status)).EqualTo(expected.status)
                  .And.Property(nameof(ApiResponse.result)).EqualTo(expected.result)
                  .And.Property(nameof(ApiResponse.errorMessage)).EqualTo(expected.errorMessage));
        }


        [Test]
        public void ConvertColumnNameToLetter_Null_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ApiExcelConverterAdapter(_excelConverter);
            ApiResponse actual = adapter.ConvertExcelColumn(null);
            ApiResponse expected = new ApiResponse();
            expected.statusCode = 422;
            expected.status = "FAILURE";
            expected.result = null;
            expected.errorMessage = "Invalid column detail provided.";
            Assert.That(actual, Has.Property(nameof(ApiResponse.statusCode)).EqualTo(expected.statusCode)
                    .And.Property(nameof(ApiResponse.status)).EqualTo(expected.status)
                    .And.Property(nameof(ApiResponse.result)).EqualTo(expected.result)
                    .And.Property(nameof(ApiResponse.errorMessage)).EqualTo(expected.errorMessage));
        }


        [Test]
        public void ConvertColumnLetterToName_Number_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ApiExcelConverterAdapter(_excelConverter);
            ColumnDetail columnDetail = new ColumnDetail();
            columnDetail.columnLetter = "55";
            ApiResponse actual = adapter.ConvertExcelColumn(columnDetail);
            ApiResponse expected = new ApiResponse();
            expected.statusCode = 422;
            expected.status = "FAILURE";
            expected.result = null;
            expected.errorMessage = "Invalid column detail provided.";
            Assert.That(actual, Has.Property(nameof(ApiResponse.statusCode)).EqualTo(expected.statusCode)
                  .And.Property(nameof(ApiResponse.status)).EqualTo(expected.status)
                  .And.Property(nameof(ApiResponse.result)).EqualTo(expected.result)
                  .And.Property(nameof(ApiResponse.errorMessage)).EqualTo(expected.errorMessage));
        }

        [Test]
        public void ConvertColumnNameToLetter_Alpha_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ApiExcelConverterAdapter(_excelConverter);
            ColumnDetail columnDetail = new ColumnDetail();
            columnDetail.columnName = "ABC";
            ApiResponse actual = adapter.ConvertExcelColumn(columnDetail);
            ApiResponse expected = new ApiResponse();
            expected.statusCode = 422;
            expected.status = "FAILURE";
            expected.result = null;
            expected.errorMessage = "Invalid column detail provided.";
            Assert.That(actual, Has.Property(nameof(ApiResponse.statusCode)).EqualTo(expected.statusCode)
                   .And.Property(nameof(ApiResponse.status)).EqualTo(expected.status)
                   .And.Property(nameof(ApiResponse.result)).EqualTo(expected.result)
                   .And.Property(nameof(ApiResponse.errorMessage)).EqualTo(expected.errorMessage));
        }

        [Test]
        public void ConvertColumnNameToLetter_InvalidInput0_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ApiExcelConverterAdapter(_excelConverter);
            ColumnDetail columnDetail = new ColumnDetail();
            columnDetail.columnName = "0";
            ApiResponse actual = adapter.ConvertExcelColumn(columnDetail);
            ApiResponse expected = new ApiResponse();
            expected.statusCode = 422;
            expected.status = "FAILURE";
            expected.result = null;
            expected.errorMessage = "Invalid column detail provided.";
            Assert.That(actual, Has.Property(nameof(ApiResponse.statusCode)).EqualTo(expected.statusCode)
                  .And.Property(nameof(ApiResponse.status)).EqualTo(expected.status)
                  .And.Property(nameof(ApiResponse.result)).EqualTo(expected.result)
                  .And.Property(nameof(ApiResponse.errorMessage)).EqualTo(expected.errorMessage));

        }


        public void ConvertColumnNameToLetter_InvalidInputminus1_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ApiExcelConverterAdapter(_excelConverter);
            ColumnDetail columnDetail = new ColumnDetail();
            columnDetail.columnName = "-1";
            ApiResponse actual = adapter.ConvertExcelColumn(columnDetail);
            ApiResponse expected = new ApiResponse();
            expected.statusCode = 422;
            expected.status = "FAILURE";
            expected.result = null;
            expected.errorMessage = "Invalid column detail provided.";
            Assert.That(actual, Has.Property(nameof(ApiResponse.statusCode)).EqualTo(expected.statusCode)
                  .And.Property(nameof(ApiResponse.status)).EqualTo(expected.status)
                  .And.Property(nameof(ApiResponse.result)).EqualTo(expected.result)
                  .And.Property(nameof(ApiResponse.errorMessage)).EqualTo(expected.errorMessage));

        }


        [Test]
        public void ConvertColumnLetterToName_InvalidInput0_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ApiExcelConverterAdapter(_excelConverter);
            ColumnDetail columnDetail = new ColumnDetail();
            columnDetail.columnLetter = "0";
            ApiResponse actual = adapter.ConvertExcelColumn(columnDetail);
            ApiResponse expected = new ApiResponse();
            expected.statusCode = 422;
            expected.status = "FAILURE";
            expected.result = null;
            expected.errorMessage = "Invalid column detail provided.";
            Assert.That(actual, Has.Property(nameof(ApiResponse.statusCode)).EqualTo(expected.statusCode)
                  .And.Property(nameof(ApiResponse.status)).EqualTo(expected.status)
                  .And.Property(nameof(ApiResponse.result)).EqualTo(expected.result)
                  .And.Property(nameof(ApiResponse.errorMessage)).EqualTo(expected.errorMessage));

        }


        public void ConvertColumnLetterToName_InvalidInputminus1_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ApiExcelConverterAdapter(_excelConverter);
            ColumnDetail columnDetail = new ColumnDetail();
            columnDetail.columnLetter = "-1";
            ApiResponse actual = adapter.ConvertExcelColumn(columnDetail);
            ApiResponse expected = new ApiResponse();
            expected.statusCode = 422;
            expected.status = "FAILURE";
            expected.result = null;
            expected.errorMessage = "Invalid column detail provided.";
            Assert.That(actual, Has.Property(nameof(ApiResponse.statusCode)).EqualTo(expected.statusCode)
                  .And.Property(nameof(ApiResponse.status)).EqualTo(expected.status)
                  .And.Property(nameof(ApiResponse.result)).EqualTo(expected.result)
                  .And.Property(nameof(ApiResponse.errorMessage)).EqualTo(expected.errorMessage));
        }


    }
}