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

    public static void PrintSearchResults(
        List<(string Month, string Name, string Payment, string Date, string Amount, string Installment)> results,
        string searchTerm)
    {
        if (results.Count == 0)
        {
            Console.WriteLine($"\nNenhum resultado encontrado para \"{searchTerm}\".");
            return;
        }

        Console.WriteLine($"\n{results.Count} resultado(s) encontrado(s) para \"{searchTerm}\":\n");

        var total = results.Sum(r => decimal.TryParse(
            r.Amount.Replace("R$", ""),
            System.Globalization.NumberStyles.Number,
            System.Globalization.CultureInfo.InvariantCulture,
            out var v) ? v : 0m);
        var totalStr = $"R${total:F2}";

        string[] headers = ["Mês", "Descrição", "Pagamento", "Data", "Valor", "Parcela"];
        int[] widths =
        [
            Math.Max(headers[0].Length, results.Max(r => r.Month.Length)),
            Math.Max(headers[1].Length, results.Max(r => r.Name.Length)),
            Math.Max(headers[2].Length, results.Max(r => r.Payment.Length)),
            Math.Max(headers[3].Length, results.Max(r => r.Date.Length)),
            Math.Max(headers[4].Length, Math.Max(results.Max(r => r.Amount.Length), totalStr.Length)),
            Math.Max(headers[5].Length, results.Max(r => r.Installment.Length))
        ];

        var separator = "─" + string.Join("─┼─", widths.Select(w => new string('─', w))) + "─";
        var headerLine = " " + string.Join(" │ ", headers.Select((h, i) => h.PadRight(widths[i]))) + " ";

        Console.WriteLine(headerLine);
        Console.WriteLine(separator);

        foreach (var r in results)
        {
            Console.WriteLine(" " + string.Join(" │ ", new[]
            {
                r.Month.PadRight(widths[0]),
                r.Name.PadRight(widths[1]),
                r.Payment.PadRight(widths[2]),
                r.Date.PadRight(widths[3]),
                r.Amount.PadRight(widths[4]),
                r.Installment.PadRight(widths[5])
            }) + " ");
        }

        Console.WriteLine(separator);
        Console.WriteLine(" " + string.Join(" │ ", new[]
        {
            "Total".PadRight(widths[0]),
            "".PadRight(widths[1]),
            "".PadRight(widths[2]),
            "".PadRight(widths[3]),
            totalStr.PadRight(widths[4]),
            "".PadRight(widths[5])
        }) + " ");
    }
}