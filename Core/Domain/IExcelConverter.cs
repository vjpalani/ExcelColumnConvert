namespace ExcelConverter.Core
{
    public interface IExcelConverter
    {
        string ConvertColumnLetterToName(string columnLetter);
        string ConvertColumnNameToLetter(string columnName);
    }
}