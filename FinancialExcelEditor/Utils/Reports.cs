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
    
    public static void PrintReport(List<(string, string)> items, string? worksheetName)
    {
        decimal total = 0;
        
        Console.WriteLine($"\nGastos - {worksheetName} de {DateTime.Now.Year}:\n");
        
        foreach (var item in items)
        {
            if (decimal.TryParse(item.Item2, out var value))
            {
                total += Math.Round(value, 2);
                Console.WriteLine($"{item.Item1}: R${value:F2}");
            }
            else
            {
                Console.WriteLine($"Valor inválido: item {item.Item2} com valor {item.Item1}");
            }
        }
        Console.WriteLine($"\nTotal: R${total:F2}");
    }
}