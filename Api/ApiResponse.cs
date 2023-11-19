using Microsoft.AspNetCore.SignalR;

namespace ExcelConverter.Model
{

  public class ApiResponse
  {
    public int statusCode { get; set; }
    public string status { get; set; }
    public string result { get; set; }
    public string errorMessage { get; set; }
  }


}