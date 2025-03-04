using ClosedXML.Excel;

namespace FinancialExcelEditor.Services;

public interface IExcelService
{
    List<(int Id, string Name)> GetWorksheetNames();
    IXLWorksheet GetWorksheet(string name);
    IXLWorksheet DuplicateWorksheet(IXLWorksheet worksheet, string newName);
    void DeleteRow(IXLWorksheet worksheet, List<int> rowNumber);
    void ClearCreditCarRow(IXLRow row);
    void SaveWorkbook();
}