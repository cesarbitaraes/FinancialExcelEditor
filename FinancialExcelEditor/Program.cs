using ClosedXML.Excel;
using FinancialExcelEditor.Controllers;
using FinancialExcelEditor.Services;
using FinancialExcelEditor.Utils;

try
{
    const string filePath = "Controle Financeiro.xlsx";

    using var workbook = new XLWorkbook(filePath);

    var excelService = new ExcelService(workbook);
    var sheetControler = new SheetController(excelService);
    var inputHelper = new InputHelper();
    var rowController = new RowController(excelService);

    var selectedTab = sheetControler.GetSelectedWorksheet(inputHelper);
    if (selectedTab == null) return;

    var newTabCreated = sheetControler.DuplicateWorksheet(selectedTab, inputHelper);

    Console.WriteLine($"Aba '{newTabCreated}' criada com sucesso.");
    rowController.ProcessRows(newTabCreated);
    excelService.SaveWorkbook();

}
catch (Exception ex)
{
    Console.WriteLine($"Ocorreu um erro na execução do programa: {ex.Message}");
}




// ToDo: 02 - Gerar relatório
// ToDo: 03 - Git