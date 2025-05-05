using ClosedXML.Excel;
using FinancialExcelEditor.Controllers;
using FinancialExcelEditor.Services;

namespace FinancialExcelEditor.Utils;

public static class WorkbookHelper
{
    public static (IXLWorksheet?, IExcelService) LoadWorkbookAndSelectSheet()
    {
        const string filePath = "Controle Financeiro.xlsx";
        var workbook = new XLWorkbook(filePath);
        IExcelService excelService = new ExcelService(workbook);
        var sheetController = new SheetController(excelService);

        var selectedTab = sheetController.GetSelectedWorksheet();
        return selectedTab != null ? (selectedTab, excelService) : (null, excelService);
    }
}