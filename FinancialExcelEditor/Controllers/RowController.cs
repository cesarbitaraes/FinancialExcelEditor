using ClosedXML.Excel;
using FinancialExcelEditor.Services;
using FinancialExcelEditor.Utils;

namespace FinancialExcelEditor.Controllers;

public class RowController(ExcelService excelService)
{
    public void ProcessRows(IXLWorksheet worksheet)
    {
        List<int> rowsToDelete = [];
        List<string> installmentOperations = [];
        
        foreach (var row in worksheet.RowsUsed())
        {
            var cellColumnFourValue = row.Cell(4).Value;
            var cellColumnSixValue = row.Cell(6).Value;
            if (cellColumnFourValue.IsDateTime && cellColumnSixValue.ToString().Equals("1"))
            {
                var cellDateValue = cellColumnFourValue.GetDateTime();
                var newDate = DateHelper.AddMonths(cellDateValue, 1);
                row.Cell(4).Value = newDate;
                excelService.ClearCreditCarRow(row);
            }
            else
            {
                var numbers = cellColumnSixValue.ToString().Split("/");
                if (int.TryParse(numbers[0], out var currentInstallment) &&
                    int.TryParse(numbers[1], out var totalInstallment))
                {
                    if (currentInstallment < totalInstallment)
                    {
                        var newInstallmentValue = currentInstallment + 1;
                        installmentOperations.Add($"Compra '{row.Cell(1).Value}' teve sua parcela alterada de {currentInstallment} para {newInstallmentValue}/{totalInstallment}.");
                        row.Cell(6).Value = $"{newInstallmentValue}/{totalInstallment}";
                    }
                    else
                    {
                        installmentOperations.Add($"Compra '{row.Cell(1).Value}' teve seu registro excluído pois foi finalizada: {currentInstallment}/{totalInstallment}.");
                        rowsToDelete.Add(row.RowNumber());
                    }
                }
            }
        }
        if (rowsToDelete.Count != 0) excelService.DeleteRow(worksheet, rowsToDelete);
        Reports.PrintInstallmentOperations(installmentOperations);
    }
}