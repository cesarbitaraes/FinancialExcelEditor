using FinancialExcelEditor.Controllers;
using FinancialExcelEditor.Services;

namespace FinancialExcelEditor.Actions;

public class GenerateDuplicateWorksheetAction(SheetController sheetController, IExcelService excelService) : IAction
{
    public void Execute()
    {
        var selectedTab = sheetController.GetSelectedWorksheet();

        if (selectedTab == null) return;

        var newTabCreated = sheetController.DuplicateWorksheet(selectedTab);

        Console.WriteLine($"Aba '{newTabCreated.Name}' criada com sucesso.");
        RowController.ProcessRows(newTabCreated, excelService);
        excelService.SaveWorkbook();
    }
}
