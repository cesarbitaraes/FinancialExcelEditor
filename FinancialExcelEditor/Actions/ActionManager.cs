namespace FinancialExcelEditor.Actions;

public class ActionManager
{
    private readonly Dictionary<string, IAction> _actions = new();

    public ActionManager()
    {
        _actions.Add("1", new GenerateDuplicateWorksheetAction());
        _actions.Add("2", new GenerateMonthlyReportAction());
        //_actions.Add("3", new SearchPurchaseAction());
    }

    public void ExecuteAction(string actionKey)
    {
        if (_actions.TryGetValue(actionKey, out IAction? value))
        {
            value.Execute();
        }
        else
        {
            Console.WriteLine("Ação inválida, por favor escolha uma ação válida.");
        }
    }
}