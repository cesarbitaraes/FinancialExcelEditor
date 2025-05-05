using FinancialExcelEditor.Controllers;
using FinancialExcelEditor.Utils;

namespace FinancialExcelEditor.Actions;

public class GenerateDuplicateWorksheetAction : IAction
{
    public void Execute()
    {
        var (selectedTab, excelService) = WorkbookHelper.LoadWorkbookAndSelectSheet();

        if (selectedTab == null) return;

        var sheetControler = new SheetController(excelService);
        var newTabCreated = sheetControler.DuplicateWorksheet(selectedTab);

        Console.WriteLine($"Aba '{newTabCreated}' criada com sucesso.");
        RowController.ProcessRows(newTabCreated, excelService);
        excelService.SaveWorkbook();
    }
}