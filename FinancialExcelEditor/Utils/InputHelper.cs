namespace FinancialExcelEditor.Utils;

public class InputHelper
{
    public string? GetValidInput(string message)
    {
        Console.WriteLine(message);
        return Console.ReadLine();
    }
}