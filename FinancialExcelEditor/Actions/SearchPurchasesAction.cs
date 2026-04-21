using FinancialExcelEditor.Controllers;
using FinancialExcelEditor.Services;
using FinancialExcelEditor.Utils;

namespace FinancialExcelEditor.Actions;

public class SearchPurchasesAction(IExcelService excelService) : IAction
{
    public void Execute()
    {
        var searchTerm = InputHelper.GetSearchInput(
            "Digite o termo de busca (pressione ESPAÇO para sair): ");

        while (searchTerm != null)
        {
            if (string.IsNullOrWhiteSpace(searchTerm))
            {
                Console.WriteLine("É necessário informar um termo de busca.");
            }
            else
            {
                var worksheets = excelService.GetMonthlyWorksheets();
                var results = RowController.SearchPurchases(worksheets, searchTerm);
                Reports.PrintSearchResults(results, searchTerm);
            }

            searchTerm = InputHelper.GetSearchInput(
                "\nDigite um novo termo para buscar novamente (ESPAÇO para sair): ");
        }

        Console.WriteLine("Busca encerrada.");
    }
}
