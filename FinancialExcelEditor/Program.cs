using FinancialExcelEditor.Actions;
using FinancialExcelEditor.Utils;

try
{
    var actionManager = new ActionManager();
    
    Console.WriteLine("Bem-vindo ao Financial Excel Editor\n");
    Console.WriteLine("O que gostaria de executar:\n");
    
    var actionChosen = InputHelper.GetValidInput(
        "1 - Gerar uma nova aba para um novo mês;\n" +
                "2 - Gerar o relatório de gastos de um determinado mês;\n" +
                "3 - Buscar alguma compra já realizada nesse ano.");
    
    if (string.IsNullOrWhiteSpace(actionChosen))
    {
        Console.WriteLine("É necessário informar uma ação. Encerrando o programa.");
        return;
    }

    actionManager.ExecuteAction(actionChosen);
}
catch (Exception ex)
{
    Console.WriteLine($"Ocorreu um erro na execução do programa: {ex.Message}");
}