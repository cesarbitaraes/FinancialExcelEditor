using ClosedXML.Excel;
using FinancialExcelEditor.Services;
using FinancialExcelEditor.Utils;

namespace FinancialExcelEditor.Controllers;

public class SheetController(IExcelService excelService)
{
    public IXLWorksheet? GetSelectedWorksheet()
    {
        Console.WriteLine("Abas disponíveis na planilha:");
        var workSheetsNames = excelService.GetWorksheetNames();
        
        foreach (var workSheetsName in workSheetsNames.Where(workSheetsName => DateHelper.CheckPortugueseMonths(workSheetsName.Name)))
        {
            Console.WriteLine(workSheetsName);
        }

        var excelTabChosen = InputHelper.GetValidInput("Qual aba gostaria de selecionar?");
        if (!string.IsNullOrWhiteSpace(excelTabChosen)) return excelService.GetWorksheet(excelTabChosen);
        Console.WriteLine("É necessário informar uma aba. Encerrando o programa.");
        return null;
    }
    
    public IXLWorksheet DuplicateWorksheet(IXLWorksheet selectedTab)
    {
        var newTabName = InputHelper.GetValidInput("Qual o nome da nova aba?");
        if (string.IsNullOrWhiteSpace(newTabName))
        {
            newTabName = $"Cópia de {selectedTab.Name}";
        }
        return excelService.DuplicateWorksheet(selectedTab, newTabName);
    }
}