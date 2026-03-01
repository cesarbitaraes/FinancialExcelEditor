using ClosedXML.Excel;
using FinancialExcelEditor.Services;
using FinancialExcelEditor.Utils;

namespace FinancialExcelEditor.Controllers;

public static class RowController
{
    public static void ProcessRows(IXLWorksheet worksheet, IExcelService excelService)
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
                var newDate = cellDateValue.AddMonths(1);
                row.Cell(4).Value = newDate;
                excelService.ClearCreditCardRow(row);
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
                        installmentOperations.Add(
                            $"Compra '{row.Cell(1).Value}' teve sua parcela alterada de {currentInstallment} para {newInstallmentValue}/{totalInstallment}.");
                        row.Cell(6).Value = $"{newInstallmentValue}/{totalInstallment}";
                    }
                    else
                    {
                        installmentOperations.Add(
                            $"Compra '{row.Cell(1).Value}' teve seu registro excluído pois foi finalizada: {currentInstallment}/{totalInstallment}.");
                        rowsToDelete.Add(row.RowNumber());
                    }
                }
            }
        }

        if (rowsToDelete.Count != 0) excelService.DeleteRow(worksheet, rowsToDelete);
        Reports.PrintInstallmentOperations(installmentOperations);
    }

    public static void CollectRowsForReport(IXLWorksheet worksheet)
    {
        var transfAnaItems = worksheet.RowsUsed()
            .Where(row => row.Cell(2).Value.ToString().Equals("Transf. Ana"))
            .Select(row => (Name: row.Cell(1).Value.ToString(), Amount: row.Cell(5).Value.ToString()))
            .ToList();

        var blackCardValue = worksheet.Cell(15, 10).Value.ToString();
        var platinumCardValue = worksheet.Cell(16, 10).Value.ToString();

        transfAnaItems.Add((Name: "Itaú Black", Amount: blackCardValue));
        transfAnaItems.Add((Name: "Itaú Platinum", Amount: platinumCardValue));

        Reports.PrintReport(transfAnaItems, worksheet.Name);
    }
}
