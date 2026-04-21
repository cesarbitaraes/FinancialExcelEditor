using ClosedXML.Excel;

namespace FinancialExcelEditor.Services;

public interface IExcelService : IDisposable
{
    List<(int Id, string Name)> GetWorksheetNames();
    IXLWorksheet GetWorksheet(string name);
    List<IXLWorksheet> GetMonthlyWorksheets();
    IXLWorksheet DuplicateWorksheet(IXLWorksheet worksheet, string newName);
    void DeleteRow(IXLWorksheet worksheet, List<int> rowNumber);
    void ClearCreditCardRow(IXLRow row);
    void SaveWorkbook();
}