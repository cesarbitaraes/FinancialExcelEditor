using FinancialExcelEditor.Controllers;

namespace FinancialExcelEditor.Actions;

public class GenerateMonthlyReportAction(SheetController sheetController) : IAction
{
    public void Execute()
    {
        var selectedTab = sheetController.GetSelectedWorksheet();

        if (selectedTab == null) return;

        RowController.CollectRowsForReport(selectedTab);
    }
}
