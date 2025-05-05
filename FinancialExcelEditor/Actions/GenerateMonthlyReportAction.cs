using FinancialExcelEditor.Controllers;
using FinancialExcelEditor.Utils;

namespace FinancialExcelEditor.Actions;

public class GenerateMonthlyReportAction : IAction
{
    public void Execute()
    {
        var (selectedTab, excelService) = WorkbookHelper.LoadWorkbookAndSelectSheet();

        if (selectedTab == null) return;
        
        RowController.CatchRollsToReport(selectedTab);
    }
}