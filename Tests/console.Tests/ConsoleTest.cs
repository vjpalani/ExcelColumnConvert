using NUnit.Framework;
using ExcelConverter.Adapters;
using ExcelConverter.Core;

namespace ExcelConverter.Tests
{
    [TestFixture]
    public class ConsoleTest
    {
        private IExcelConverter _excelConverter;

        [SetUp]
        public void Setup()
        {
            _excelConverter = new ExcelConverterClass();
        }

        [Test]
        public void ConvertColumnLetterToName_ValidInput_ShouldReturnCorrectName()
        {
            IExcelConverterAdapter adapter = new ConsoleExcelConverterAdapter(_excelConverter);
            string columnLetter = "AB";
            string expected = "The numerical representation corresponding to the letter 'AB' in Excel is  28.";
            string actual = adapter.ConvertColumnLetterToName(columnLetter);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConvertColumnNameToLetter_ValidInput_ShouldReturnCorrectLetter()
        {
            IExcelConverterAdapter adapter = new ConsoleExcelConverterAdapter(_excelConverter);
            string columnName = "28";
            string expected = "The letter corresponding to the numerical value 28 in Excel is  'AB'.";
            string actual = adapter.ConvertColumnNameToLetter(columnName);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConvertColumnNameToLetter_Null_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ConsoleExcelConverterAdapter(_excelConverter);
            string actual = adapter.ConvertColumnNameToLetter(null);
            string expected = "Column name cannot be null or empty.";
            Assert.AreEqual(expected, actual);
        }
        [Test]
        public void ConvertColumnNameToLetter_isEmpty_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ConsoleExcelConverterAdapter(_excelConverter);
            string actual = adapter.ConvertColumnNameToLetter(string.Empty);
            string expected = "Column name cannot be null or empty.";
            Assert.AreEqual(expected, actual);

        }

        [Test]
        public void ConvertColumnLetterToName_Null_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ConsoleExcelConverterAdapter(_excelConverter);
            string actual = adapter.ConvertColumnLetterToName(null);
            string expected = "Column letter cannot be null or empty.";
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConvertColumnLetterToName_isEmpty_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ConsoleExcelConverterAdapter(_excelConverter);
            string actual = adapter.ConvertColumnLetterToName(string.Empty);
            string expected = "Column letter cannot be null or empty.";
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConvertColumnNameToLetter_InvalidInput0_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ConsoleExcelConverterAdapter(_excelConverter);
            string expected = "Column name cannot be 0.";
            string actual = adapter.ConvertColumnNameToLetter("0");
            Assert.AreEqual(expected, actual);
        }


        public void ConvertColumnNameToLetter_InvalidInputminus1_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ConsoleExcelConverterAdapter(_excelConverter);
            string expected = "Invalid column name.";
            string actual = adapter.ConvertColumnNameToLetter("-1");
            Assert.AreEqual(expected, actual);
        }

        public void ConvertColumnNameToLetter_InvalidInputalpha_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ConsoleExcelConverterAdapter(_excelConverter);
            string expected = "Invalid column name.";
            string actual = adapter.ConvertColumnNameToLetter("ABC");
            Assert.AreEqual(expected, actual);
        }


        [Test]
        public void ConvertColumnLetterToName_InvalidInput0_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ConsoleExcelConverterAdapter(_excelConverter);
            string expected = "Column letter cannot be 0.";
            string actual = adapter.ConvertColumnLetterToName("0");
            Assert.AreEqual(expected, actual);
        }


        public void ConvertColumnLetterToName_InvalidInputminus1_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ConsoleExcelConverterAdapter(_excelConverter);
            string expected = "Invalid column letter.";
            string actual = adapter.ConvertColumnLetterToName("-1");
            Assert.AreEqual(expected, actual);
        }

        public void ConvertColumnLetterToName_InvalidInputalpha_ThrowsErrorMessage()
        {
            IExcelConverterAdapter adapter = new ConsoleExcelConverterAdapter(_excelConverter);
            string expected = "Invalid column letter.";
            string actual = adapter.ConvertColumnLetterToName("ABC");
            Assert.AreEqual(expected, actual);
        }
    }

}