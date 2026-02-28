namespace FinancialExcelEditor.Utils;

public static class InputHelper
{
    public static string? GetValidInput(string message)
    {
        Console.WriteLine(message);
        return Console.ReadLine();
    }

    public static void FinalMessage()
    {
        Console.WriteLine("\nPressione qualquer tecla para encerrar o programa.");
        Console.ReadKey(intercept: true);
    }
}