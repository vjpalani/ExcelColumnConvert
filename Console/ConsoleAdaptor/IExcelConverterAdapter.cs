namespace ExcelConverter.Adapters
{
    public interface IExcelConverterAdapter
    {
        string ConvertColumnLetterToName(string columnLetter);
        string ConvertColumnNameToLetter(string columnName);
    }
}