namespace FinancialExcelEditor.Utils;

public static class Reports
{
    public static void PrintInstallmentOperations(List<string> operations)
    {
        Console.WriteLine("As seguintes alterações em parcelas foram executadas:");

        foreach (var operation in operations)
        {
            Console.WriteLine(operation);
        }
    }
}