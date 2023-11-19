using NUnit.Framework;
using ExcelConverter.Core;

namespace ExcelConverter.Tests
{
    [TestFixture]
    public class CoreTest
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

            string columnLetter = "AB";
            string expectedName = "28";

            string actualName = _excelConverter.ConvertColumnLetterToName(columnLetter);

            Assert.AreEqual(expectedName, actualName);
        }

        [Test]
        public void ConvertColumnNameToLetter_ValidInput_ReturnsCorrectLetter()
        {
            string columnName = "28";
            string expectedLetter = "AB";
            string actualLetter = _excelConverter.ConvertColumnNameToLetter(columnName);
            Assert.AreEqual(expectedLetter, actualLetter);
        }

        [Test]
        public void ConvertColumnLetterToName_NullOrEmptyInput_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>(() => _excelConverter.ConvertColumnLetterToName(null));
            Assert.Throws<ArgumentException>(() => _excelConverter.ConvertColumnLetterToName(string.Empty));
        }

        [Test]
        public void ConvertColumnNameToLetter_NullOrEmptyInput_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>(() => _excelConverter.ConvertColumnNameToLetter(null));
            Assert.Throws<ArgumentException>(() => _excelConverter.ConvertColumnNameToLetter(string.Empty));
        }

        public void ConvertLetterToName_InvalidInput_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>(() => _excelConverter.ConvertColumnLetterToName("0"));
            Assert.Throws<ArgumentException>(() => _excelConverter.ConvertColumnLetterToName("-1"));
            Assert.Throws<ArgumentException>(() => _excelConverter.ConvertColumnLetterToName("5"));
            Assert.Throws<ArgumentException>(() => _excelConverter.ConvertColumnLetterToName("S5"));

        }

        [Test]
        public void ConvertColumnNameToLetter_InvalidInput_ThrowsArgumentException()
        {
            Assert.Throws<ArgumentException>(() => _excelConverter.ConvertColumnNameToLetter("0"));
            Assert.Throws<ArgumentException>(() => _excelConverter.ConvertColumnNameToLetter("-1"));
            Assert.Throws<ArgumentException>(() => _excelConverter.ConvertColumnNameToLetter("ABC"));
            Assert.Throws<ArgumentException>(() => _excelConverter.ConvertColumnNameToLetter("5BC"));

        }
    }
}