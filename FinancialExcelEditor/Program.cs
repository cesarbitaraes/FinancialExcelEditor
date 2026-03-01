using FinancialExcelEditor.Actions;
using FinancialExcelEditor.Controllers;
using FinancialExcelEditor.Services;
using FinancialExcelEditor.Utils;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

try
{
    var configuration = new ConfigurationBuilder()
        .AddJsonFile("appsettings.json", optional: false)
        .Build();

    var settings = new AppSettings();
    configuration.Bind(settings);

    using var serviceProvider = new ServiceCollection()
        .AddSingleton(settings)
        .AddSingleton<IExcelService, ExcelService>()
        .AddTransient<SheetController>()
        .AddTransient<GenerateDuplicateWorksheetAction>()
        .AddTransient<GenerateMonthlyReportAction>()
        .AddTransient<ActionManager>()
        .BuildServiceProvider();

    var actionManager = serviceProvider.GetRequiredService<ActionManager>();

    Console.WriteLine("Bem-vindo ao Financial Excel Editor\n");
    Console.WriteLine("Escolha a opção que deseja executar:\n");

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
finally
{
    InputHelper.FinalMessage();
}
