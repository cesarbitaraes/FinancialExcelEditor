namespace FinancialExcelEditor.Utils;

public static class InputHelper
{
    public static string? GetValidInput(string message)
    {
        Console.WriteLine(message);
        return Console.ReadLine();
    }

    /// <summary>
    /// Lê um termo de busca caractere a caractere.
    /// Espaços são ignorados silenciosamente (garante entrada de uma só palavra).
    /// Retorna null se o usuário pressionar ESC.
    /// </summary>
    public static string? GetSearchInput(string message)
    {
        Console.Write(message);
        var input = new System.Text.StringBuilder();

        while (true)
        {
            var key = Console.ReadKey(intercept: true);

            if (key.Key == ConsoleKey.Spacebar)
            {
                Console.WriteLine();
                return null;
            }

            if (key.Key == ConsoleKey.Enter)
            {
                Console.WriteLine();
                return input.ToString();
            }

            if (key.Key == ConsoleKey.Backspace && input.Length > 0)
            {
                input.Remove(input.Length - 1, 1);
                Console.Write("\b \b");
                continue;
            }

            if (!char.IsControl(key.KeyChar))
            {
                input.Append(key.KeyChar);
                Console.Write(key.KeyChar);
            }
        }
    }

    public static void FinalMessage()
    {
        Console.WriteLine("\nPressione qualquer tecla para encerrar o programa.");
        Console.ReadKey(intercept: true);
    }
}