using ClosedXML.Excel;
using FinancialExcelEditor.Controllers;
using FinancialExcelEditor.Services;
using FinancialExcelEditor.Utils;

try
{
    const string filePath = "Controle Financeiro.xlsx";

    using var workbook = new XLWorkbook(filePath);

    IExcelService excelService = new ExcelService(workbook);
    var sheetControler = new SheetController(excelService);
    var inputHelper = new InputHelper();

    var selectedTab = sheetControler.GetSelectedWorksheet(inputHelper);
    if (selectedTab == null) return;

    var newTabCreated = sheetControler.DuplicateWorksheet(selectedTab, inputHelper);

    Console.WriteLine($"Aba '{newTabCreated}' criada com sucesso.");
    RowController.ProcessRows(newTabCreated, excelService);
    excelService.SaveWorkbook();

}
catch (Exception ex)
{
    Console.WriteLine($"Ocorreu um erro na execução do programa: {ex.Message}");
}

// ToDo: 01 - Gerar relatório