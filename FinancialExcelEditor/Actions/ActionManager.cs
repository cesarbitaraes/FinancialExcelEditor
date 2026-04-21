namespace FinancialExcelEditor.Actions;

public class ActionManager(GenerateDuplicateWorksheetAction duplicateWorksheetAction, GenerateMonthlyReportAction monthlyReportAction, SearchPurchasesAction searchPurchasesAction)
{
    private readonly Dictionary<string, IAction> _actions = new()
    {
        { "1", duplicateWorksheetAction },
        { "2", monthlyReportAction },
        { "3", searchPurchasesAction }
    };

    public void ExecuteAction(string actionKey)
    {
        if (_actions.TryGetValue(actionKey, out IAction? value))
        {
            value.Execute();
        }
        else
        {
            Console.WriteLine("Ação inválida! Por favor, escolha uma ação válida.");
        }
    }
}
