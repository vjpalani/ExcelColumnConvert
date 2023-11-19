using System.Text;
namespace ExcelConverter.Core
{
    public class ExcelConverterClass : IExcelConverter
    {
        // Converts a column letter to its corresponding column name.
        //
        // Parameters:
        //   columnLetter: The column letter to convert.
        //
        // Returns:
        //   The corresponding column name as a string.

        public string ConvertColumnLetterToName(string columnLetter)
        {
            // Validate input
            if (string.IsNullOrEmpty(columnLetter))
                throw new ArgumentException("Column letter cannot be null or empty.");

            if (columnLetter == "0")
                throw new ArgumentException("Column letter cannot be 0.");

            // Convert column letter to column name
            int columnNumber = 0;
            foreach (char c in columnLetter)
            {
                columnNumber *= 26;
                columnNumber += c - 'A' + 1;
            }

            return columnNumber.ToString();
        }

        // Converts a column name to its corresponding column letter.
        // 
        // Parameters:
        //   columnName: The name of the column to convert.
        //
        // Returns:
        //   A string representing the column letter.


        public string ConvertColumnNameToLetter(string columnName)
        {
            // Validate input
            if (string.IsNullOrEmpty(columnName))
                throw new ArgumentException("Column name cannot be null or empty.");

            if (columnName == "0")
                throw new ArgumentException("Column name cannot be 0.");

            // Convert column name to column letter
            int columnNumber;
            if (!int.TryParse(columnName, out columnNumber) || columnNumber <= 0)
                throw new ArgumentException("Invalid column name.");

            StringBuilder columnLetter = new StringBuilder();
            while (columnNumber > 0)
            {
                int remainder = (columnNumber - 1) % 26;
                columnLetter.Insert(0, (char)('A' + remainder));
                columnNumber = (columnNumber - 1) / 26;
            }

            return columnLetter.ToString();
        }
    }
}