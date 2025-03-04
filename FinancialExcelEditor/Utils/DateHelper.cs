namespace FinancialExcelEditor.Utils;

public class DateHelper
{
    public static DateTime AddMonths(DateTime date, int months)
    {
        return date.AddMonths(months);
    }

    public static bool CheckPortugueseMonths(string portugueseMonth)
    {
        string[] months =
        [
            "janeiro", "fevereiro", "março", "abril", "maio", "junho", 
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
        ];
        
        return months.Contains(portugueseMonth, StringComparer.OrdinalIgnoreCase);
    }
}