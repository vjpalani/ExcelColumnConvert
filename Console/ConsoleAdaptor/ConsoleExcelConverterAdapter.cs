using ExcelConverter.Core;


namespace ExcelConverter.Adapters
{
    public class ConsoleExcelConverterAdapter : IExcelConverterAdapter
    {
        private readonly IExcelConverter _excelConverter;

        public ConsoleExcelConverterAdapter(IExcelConverter excelConverter)
        {
            _excelConverter = excelConverter;
        }

        // Converts a column letter to its corresponding numerical representation in Excel.
        //
        // Parameters:
        //   columnLetter:
        //     The column letter to be converted.
        //
        // Returns:
        //     The numerical representation corresponding to the letter 'columnLetter' in Excel.
        //     If an error occurs during the conversion, the error message is returned.


        public string ConvertColumnLetterToName(string columnLetter)
        {
            try
            {
                string Value = _excelConverter.ConvertColumnLetterToName(columnLetter);

                string return_object = "The numerical representation corresponding to the letter '" + columnLetter + "' in Excel is  " + Value + ".";

                return (return_object);
            }
            catch (Exception ex)
            {
                return (ex.Message);

            }

        }


        // Converts the given column name to its corresponding letter in Excel.
        public string ConvertColumnNameToLetter(string columnName)
        {
            try
            {

                string Value = _excelConverter.ConvertColumnNameToLetter(columnName);

                string return_object = "The letter corresponding to the numerical value " + columnName + " in Excel is  '" + Value + "'.";

                return (return_object);
            }
            catch (Exception ex)
            {
                return (ex.Message);

            }


        }
    }
}