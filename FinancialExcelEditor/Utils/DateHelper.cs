namespace FinancialExcelEditor.Utils;

public static class DateHelper
{
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