namespace FinancialExcelEditor.Utils;

public static class InputHelper
{
    public static string? GetValidInput(string message)
    {
        Console.WriteLine(message);
        return Console.ReadLine();
    }
}