# Automação de Tarefas no Google Sheets

Este projeto contém um script em Google Apps Script que automatiza várias tarefas dentro de uma planilha do Google Sheets. As automações incluem validação de filtros, tratamento de dados, formatação, e a criação de um menu personalizado para facilitar o uso dessas funções.

## Funcionalidades

1. **Menu Personalizado**: 
   - O script cria um menu personalizado chamado "Minhas Macros" na barra de menus do Google Sheets.
   - Este menu inclui as seguintes opções:
     - `INICIAR CHECAGEM`: Inicia o processo de verificação de dados.
     - `TRATAMENTO DOS DADOS`: Executa a rotina de tratamento de dados.
     - `FORMATAÇÃO`: Aplica a formatação pré-definida aos dados na planilha.

2. **Validação de Filtros**:
   - Verifica se há filtros aplicados na planilha e os remove, se necessário, para garantir que os dados sejam manipulados corretamente.

3. **Tratamento de Dados**:
   - Inclui funções para reexibir colunas ocultas, formatar dados, e outros procedimentos essenciais para a manutenção e análise correta dos dados na planilha.

4. **Formatação**:
   - Aplicação de formatação específica para garantir que os dados estejam apresentados de maneira clara e consistente.

## Uso

1. Abra o Google Sheets e vá para `Extensões > Apps Script`.
2. Cole o código do arquivo `automacoes.gs` no editor do Google Apps Script.
3. Salve o script.
4. Agora, ao abrir a planilha, você verá um menu chamado "Minhas Macros" na barra de ferramentas, que permitirá acessar as automações descritas.
