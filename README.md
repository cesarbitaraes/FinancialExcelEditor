# Financial Excel Editor

Ferramenta de linha de comando para automatizar tarefas recorrentes em planilhas de controle financeiro no formato `.xlsx`.

## Funcionalidades

| Opção | Descrição |
|-------|-----------|
| 1     | Gera uma nova aba para o mês seguinte, copiando a aba selecionada, incrementando parcelas e avançando datas de débito automático. |
| 2     | Exibe o relatório de gastos do mês selecionado, listando transferências e totais dos cartões de crédito. |

## Requisitos

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- Arquivo `Controle Financeiro.xlsx` na pasta de saída (ou caminho configurado em `appsettings.json`)

## Configuração

Edite `appsettings.json` antes de compilar (ou na pasta de saída após compilar):

```json
{
  "ExcelFilePath": "Controle Financeiro.xlsx",
  "CreditCards": ["Itaú Black", "Itaú Platinum", "Nubank"]
}
```

| Chave          | Descrição                                                        |
|----------------|------------------------------------------------------------------|
| `ExcelFilePath`| Caminho para o arquivo `.xlsx` (relativo ao executável).         |
| `CreditCards`  | Lista de nomes de cartões cujas células de pagador são limpas ao duplicar a aba. |

## Como executar

```bash
dotnet run --project FinancialExcelEditor/FinancialExcelEditor.csproj
```

Ou publique e execute o binário gerado em `bin/`:

```bash
dotnet publish -c Release
./bin/Release/net8.0/FinancialExcelEditor
```

## Estrutura do projeto

```
FinancialExcelEditor/
├── Actions/
│   ├── IAction.cs                          # Interface comum para as ações
│   ├── ActionManager.cs                    # Despacha a ação escolhida pelo usuário
│   ├── GenerateDuplicateWorksheetAction.cs # Ação 1: duplicar aba do mês
│   └── GenerateMonthlyReportAction.cs      # Ação 2: relatório mensal
├── Controllers/
│   ├── RowController.cs                    # Processa linhas da planilha
│   └── SheetController.cs                  # Seleciona e duplica abas
├── Services/
│   ├── IExcelService.cs                    # Abstração de acesso ao Excel
│   └── ExcelService.cs                     # Implementação com ClosedXML
├── Utils/
│   ├── AppSettings.cs                      # POCO de configuração
│   ├── DateHelper.cs                       # Validação de meses em português
│   ├── InputHelper.cs                      # Leitura de input do console
│   └── Reports.cs                          # Impressão de relatórios
├── appsettings.json                        # Configurações da aplicação
└── Program.cs                              # Ponto de entrada e DI container
```

## Dependências

| Pacote | Versão | Uso |
|--------|--------|-----|
| [ClosedXML](https://github.com/ClosedXML/ClosedXML) | 0.104.2 | Leitura e escrita de arquivos `.xlsx` |
| Microsoft.Extensions.DependencyInjection | 8.0.1 | Container de injeção de dependência |
| Microsoft.Extensions.Configuration.Json | 8.0.1 | Leitura do `appsettings.json` |
| Microsoft.Extensions.Configuration.Binder | 8.0.2 | Bind do JSON para o POCO `AppSettings` |
