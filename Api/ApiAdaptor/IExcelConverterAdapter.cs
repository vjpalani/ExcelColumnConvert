using ExcelConverter.Model;
using Newtonsoft.Json.Linq;

namespace ExcelConverter.Adapters
{
    public interface IExcelConverterAdapter
    {
        ApiResponse ConvertExcelColumn(ColumnDetail columnDetail);
    }
}