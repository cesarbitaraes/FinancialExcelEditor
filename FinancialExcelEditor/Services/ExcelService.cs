using ClosedXML.Excel;
using FinancialExcelEditor.Utils;

namespace FinancialExcelEditor.Services;

public class ExcelService(AppSettings settings) : IExcelService
{
    private readonly XLWorkbook _workbook = new(settings.ExcelFilePath);

    public List<(int Id, string Name)> GetWorksheetNames()
    {
        List<(int Id, string Name)> worksheets = [];
        var i = 1;
        foreach (var ws in _workbook.Worksheets)
        {
            worksheets.Add((i, ws.Name));
            i++;
        }
        return worksheets;
    }

    public IXLWorksheet GetWorksheet(string name)
    {
        return _workbook.Worksheets.Worksheet(name);
    }

    public IXLWorksheet DuplicateWorksheet(IXLWorksheet worksheet, string newName)
    {
        var newWorkSheet = worksheet.CopyTo(newName);
        newWorkSheet.ShowGridLines = false;
        return newWorkSheet;
    }

    public void DeleteRow(IXLWorksheet worksheet, List<int> rowNumber)
    {
        for (var j = rowNumber.Count - 1; j >= 0; j--)
            for (var k = 1; k <= 6; k++)
            {
                worksheet.Row(rowNumber[j]).Cell(k).Delete(XLShiftDeletedCells.ShiftCellsUp);
            }
    }

    public void ClearCreditCardRow(IXLRow row)
    {
        var creditCardUsed = row.Cell(2).Value.ToString();
        if (settings.CreditCards.Contains(creditCardUsed, StringComparer.OrdinalIgnoreCase)) row.Cell(2).Clear();
    }

    public void SaveWorkbook()
    {
        _workbook.Save();
    }

    public void Dispose() => _workbook.Dispose();
}
