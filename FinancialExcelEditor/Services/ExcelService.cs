using ClosedXML.Excel;

namespace FinancialExcelEditor.Services;

public class ExcelService(XLWorkbook workbook) : IExcelService
{
    public List<(int Id, string Name)> GetWorksheetNames()
    {
        List<(int Id, string Name)> worksheets = [];
        var i = 1;
        foreach (var ws in workbook.Worksheets)
        {
            worksheets.Add((i, ws.Name));
            i++;
        }
        return worksheets;
    }

    public IXLWorksheet GetWorksheet(string name)
    {
        return workbook.Worksheets.Worksheet(name);
    }

    public IXLWorksheet DuplicateWorksheet(IXLWorksheet worksheet, string newName)
    {
        var newWorkSheet = worksheet.CopyTo(newName);
        newWorkSheet.ShowGridLines = false;
        return newWorkSheet;
    }
    
    public void DeleteRow(IXLWorksheet worksheet, List<int> rowNumber){
        for (var j = rowNumber.Count - 1; j >= 0; j--)
            for (var k = 1; k <= 6; k++)
        {
            worksheet.Row(rowNumber[j]).Cell(k).Delete(XLShiftDeletedCells.ShiftCellsUp);
        }
    }

    public void ClearCreditCarRow(IXLRow row)
    {
        string[] creditCars = ["Itaú Black", "Itaú Platinum", "Nubank"];
        var creditCardUsed = row.Cell(2).Value.ToString();
        if (creditCars.Contains(creditCardUsed, StringComparer.OrdinalIgnoreCase)) row.Cell(2).Clear();
    }
    
    public void SaveWorkbook()
    {
        workbook.Save();
    }
}