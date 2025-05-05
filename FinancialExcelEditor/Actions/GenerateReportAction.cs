using FinancialExcelEditor.Controllers;
using FinancialExcelEditor.Utils;

namespace FinancialExcelEditor.Actions;

public class GenerateReportAction : IAction
{
    public void Execute()
    {
        var (selectedTab, excelService) = WorkbookHelper.LoadWorkbookAndSelectSheet();

        if (selectedTab == null) return;
        
        RowController.GenerateReport(selectedTab);
    }
}