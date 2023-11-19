using ExcelConverter.Core;
using ExcelConverter.Adapters;

namespace ConsoleAdaptor
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter a column name or number:");
            string input = Console.ReadLine();

            IExcelConverter excelConverter = new ExcelConverterClass();
            IExcelConverterAdapter consoleAdapter = new ConsoleExcelConverterAdapter(excelConverter);
            string result;
            if (int.TryParse(input, out int columnNumber))
            {
                Console.WriteLine($"{columnNumber}");
                result = consoleAdapter.ConvertColumnNameToLetter(columnNumber.ToString());
                Console.WriteLine($"{result}");
            }
            else
            {
                Console.WriteLine($"{input}");
                result = consoleAdapter.ConvertColumnLetterToName(input);
                Console.WriteLine($"{result}");
            }
        }
    }
}