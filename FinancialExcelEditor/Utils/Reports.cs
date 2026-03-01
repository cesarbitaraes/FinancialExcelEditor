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
    
    public static void PrintReport(List<(string Name, string Amount)> items, string worksheetName)
    {
        decimal total = 0;

        Console.WriteLine($"\nGastos - {worksheetName} de {DateTime.Now.Year}:\n");

        foreach (var item in items)
        {
            if (decimal.TryParse(item.Amount, out var value))
            {
                total += Math.Round(value, 2);
                Console.WriteLine($"{item.Name}: R${value:F2}");
            }
            else
            {
                Console.WriteLine($"Valor inválido: item {item.Name} com valor {item.Amount}");
            }
        }
        Console.WriteLine($"\nTotal: R${total:F2}");
    }
}