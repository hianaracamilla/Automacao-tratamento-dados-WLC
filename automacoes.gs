function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Minhas Macros')
      .addItem('INICIAR CHECAGEM', 'IniciarChecagem')
      .addItem('TRATAMENTO DOS DADOS', 'TratamentoDeDados')
      .addItem('FORMATAÇÃO', 'Formatacao')
      .addToUi();
}

function ValidacaoFiltros(sheet){
  // Verifica se tem filtro na planilha
  var filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }
}

// Função para reexibir colunas ocultas na página indicada
function ReexibirColunasOcultas(sheet) {
  var lastColumn = sheet.getLastColumn();

  for (var i = 1; i <= lastColumn; i++) {
    if (sheet.isColumnHiddenByUser(i)) {
      sheet.showColumns(i);
    }
  }
}

function prepararBase(sheet) {
  // Verifica se a primeira célula possui conteúdo, se não, exclui a primeira linha
  var primeiraLinha = sheet.getRange('A1').getValue();
  if (primeiraLinha == "") {
    sheet.deleteRows(1, 1);
  }

  sheet.deleteColumns(8, 1)
}

function ConfirmarEfectivo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  for (var i = 1; i <= lastRow; i++) {
    var tipo = sheet.getRange(i, 7).getValue(); // Coluna G (7)
    if (tipo === "Efectivo oficina Santiago" || tipo === "Efectivo oficina ATACAMA") {
      var valorH = sheet.getRange(i, 8).getValue(); // Coluna H (8)
      sheet.getRange(i, 10).setValue(valorH); // Coluna J (10)
    }
  }
}


function classificaBase(sheet){
  var ultimaLinha = sheet.getLastRow();

  //verifica se tem filtro disponivel e retira

  var filter = sheet.getRange(1, 1, ultimaLinha, 9).createFilter();

  // Classifica Data A a Z
  filter.sort(2, true);
  
  //Classifica cliente de A a Z
  filter.sort(3, true);

  // Classifica forma de pagamento Z a A
    filter.sort(7, false);

  filter.remove();
}

function Teste(){
    var ss = SpreadsheetApp.getActive();
  var checagem = ss.getSheetByName('CHECAGEM');

  Cabecalho(checagem)
}


function Cabecalho(sheet){

  sheet.getRange('J1').setFormula('=SUMIFS(P:P;G:G;I1;B:B;$H$1)')
  sheet.getRange('J2').setFormula('=SUMIFS(P:P;G:G;I2;B:B;$H$1)')
  sheet.getRange('J3').setFormula('=SUMIFS(P:P;G:G;I3;B:B;$H$1)')
  sheet.getRange('J4').setFormula('=COUNTIF(M:M;"SIM")')
  sheet.getRange('L4').setFormula('=(COUNTIF(J:J;"")-1)')
  sheet.getRange('K4').setValue('SEM ID').setBackground("#4BACC6")
  sheet.getRange('K4').setFontColor("#ffffff")
}


function AplicarFormatacaoCondicionalCompleta(aba) {
  
  // Limpa todas as regras de formatação condicional existentes
  aba.clearConditionalFormatRules();
  
  var ultimaLinha = aba.getLastRow();
  
  // Define as faixas onde quer aplicar a formatação condicional
  var range = aba.getRange('H7:Q' + ultimaLinha);
  var range2 = aba.getRange('J4:L4')

 
  // Define as regras de formatação condicional
  var regraH7L = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ISBLANK($H7:$L7)')
    .setBackground("#C0C0C0")
    .setRanges([range])
    .build();

  var regraH = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=MATCH("não localizado";$H7:$L7; 0)')
    .setBackground("#8b008b")
    .setFontColor("ffffff")
    .setRanges([range])
    .build();

    var regraR = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$R7="VERIFICAR"')
    .setBackground("#b7e1cd")
    .setRanges([range])
    .build();
  
  var regraQ = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($Q7>999;$Q7<-999)')
    .setBackground("#DE3163")
    .setRanges([range])
    .build();
  
  var regraS = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$S7="VERIFICAR"')
    .setBackground("#ffa500")
    .setRanges([range])
    .build();
  
  var regraT = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$T7="VERIFICAR"')
    .setBackground("#006400")
    .setRanges([range])
    .build();
  
  var regraJ = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('COUNTIF(J:J, J1)>1')
    .setBackground("#1976D2")
    .setRanges([range])
    .build();

  var regraCabecalho = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=J4 != L4')
    .setBackground("#DE3163")
    .setFontColor("ffffff")
    .setRanges([range2])
    .build();
  
  // Aplica as novas regras de formatação condicional
  var regrasAtuais = aba.getConditionalFormatRules();
  regrasAtuais.push(regraH7L,regraH, regraQ, regraR, regraS, regraT, regraJ, regraCabecalho);
  aba.setConditionalFormatRules(regrasAtuais);
}


function limparChecagem(sheet){
  var ultimaLinha = sheet.getLastRow();
  sheet.getRange(7, 1, ultimaLinha, sheet.getMaxColumns()).clear();
}

function copiarDadosBase(sheetInicial, sheetFinal){

  var extensao = sheetInicial.getLastRow() - 1;

  sheetInicial.getRange(2, 1, extensao, 7).copyTo(sheetFinal.getRange(7, 1, extensao, 7))

  sheetInicial.getRange(2, 9, extensao, 1).copyTo(sheetFinal.getRange(7, 8, extensao, 1))

  sheetInicial.getRange(2, 8, extensao, 1).copyTo(sheetFinal.getRange(7, 21, extensao, 1))

  sheetInicial.getRange(1, 1, extensao + 1, sheetInicial.getMaxColumns()).clear()

  sheetFinal.getRange('C:C').setHorizontalAlignment('left')
  sheetFinal.getRange('G:G').setHorizontalAlignment('left')
}

function ValidacaoEfectivo(sheet) {
  var range = sheet.getRange(1, 7, sheet.getLastRow(), 1); // Pega o intervalo da coluna 7
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    var efectivoValue = values[i][0];
    var lojaValue = sheet.getRange(i + 7, 5).getValue(); // Pega o valor da coluna 5 (coluna E) na mesma linha

    if (efectivoValue === "Efectivo oficina Santiago") {
      if (lojaValue === "LOJA SANTIAGO") {
        sheet.getRange(i + 7, 10).setValue("-");
      } else {
        sheet.getRange(i + 7, 10).setValue("VERIFICAR");
      }
    } else if (efectivoValue === "Efectivo oficina ATACAMA") {
      if (lojaValue === "LOJA ATACAMA") {
        sheet.getRange(i + 7, 10).setValue("-");
      } else {
        sheet.getRange(i + 7, 10).setValue("VERIFICAR");
      }
    }
  }
}

function Filtrar(sheet, coluna, criteria){
  ValidacaoFiltros(sheet); // Função que remove filtros, se existente

  // Cria filtro na planilha
  var filter = sheet.getRange(6, 1, sheet.getLastRow() - 5, sheet.getMaxColumns()).createFilter();

  filter.setColumnFilterCriteria(coluna, criteria);

}

function OcutarColunas(sheet, colInicial, colFinal){
  sheet.hideColumns(colInicial, colFinal);
}

function cruzarDados(sheetChecagem, sheetLogin) {

  var rangeChecagem = sheetChecagem.getRange(7, 4, sheetChecagem.getLastRow() - 6, 1); // Coluna 4 da linha 7 até a última linha com dados
  var valuesChecagem = rangeChecagem.getValues();

  var rangeLogin = sheetLogin.getRange(2, 1, sheetLogin.getLastRow() - 1, 2); // Colunas 1 e 2 da linha 2 até a última linha com dados
  var valuesLogin = rangeLogin.getValues();

  var loginMap = new Map();

  // Mapeia os valores da coluna 1 para os valores da coluna 2 na planilha LOGIN
  for (var i = 0; i < valuesLogin.length; i++) {
    loginMap.set(valuesLogin[i][0], valuesLogin[i][1]);
  }

  // Percorre os valores da planilha CHECAGEM, cruzando os dados e escrevendo o resultado na coluna 5
  for (var j = 0; j < valuesChecagem.length; j++) {
    var checagemValue = valuesChecagem[j][0];
    if (loginMap.has(checagemValue)) {
      sheetChecagem.getRange(j + 7, 5).setValue(loginMap.get(checagemValue)); // Coluna 5 na linha correspondente
    } else {
      sheetChecagem.getRange(j + 7, 5).setValue('Não localizado');
    }
  }
}

function ExcluirLinhasAbaixoUltimaLinha(sheet) {
  var lastRow = sheet.getLastRow();
  var maxRows = sheet.getMaxRows();
  
  if (lastRow < maxRows) {
    sheet.deleteRows(lastRow, maxRows - (lastRow - 1));
  }
}

function RemoverHashtags(sheet) {
  var startRow = 1; // Começa na primeira linha
  var column = 10; // Coluna J (10)
  var lastRow = sheet.getLastRow();
  
  // Percorre todas as linhas na coluna J
  for (var i = startRow; i <= lastRow; i++) {
    var cell = sheet.getRange(i, column);
    var cellValue = cell.getValue().toString(); // Converte o valor da célula para string
    if (cellValue.includes("#")) {
      var newValue = cellValue.replace(/#/g, ""); // Remove todas as ocorrências de "#"
      cell.setValue(newValue);
    }
  }
}

function ValidacoesErros(sheet){
  var startCol = 16; // Coluna P (16)
  var endCol = 20; // Coluna t (20)

  var lastRow = sheet.getLastRow();

  sheet.getRange('P5:t5').copyTo(sheet.getRange(7, startCol, lastRow, endCol - startCol + 1  ), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

function ConfirmarMercadoPago() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  for (var i = 1; i <= lastRow; i++) {
    var tipo = sheet.getRange(i, 7).getValue(); // Coluna G (7)
    if (tipo === "MERCADO PAGO") {
      var valorH = sheet.getRange(i, 8).getValue(); // Coluna H (8)
      var valorI = sheet.getRange(i, 9).getValue(); // Coluna I (9)
      var valorJ = sheet.getRange(i, 10).getValue(); // Coluna J (10)
      
      sheet.getRange(i, 9).setValue(valorH); // Coluna I (9)
      sheet.getRange(i, 10).setValue(valorI); // Coluna J (10)
      sheet.getRange(i, 11).setValue(valorJ); // Coluna K (11)
    }
  }
}

function valoresWise(sheet) {
  var lastRow = sheet.getLastRow();
  
  for (var i = 1; i <= lastRow; i++) {
    var tipo = sheet.getRange(i, 7).getValue(); // Coluna G (7)
    if (tipo === "WISE") {
      var valorP = sheet.getRange(i, 16).getValue(); // Coluna P (16)
      sheet.getRange(i, 9).setValue(valorP); // Coluna I (9)
    }
  }
}


function IniciarChecagem(){
  var ss = SpreadsheetApp.getActive();
  var base = ss.getSheetByName('Sheet1');
  var checagem = ss.getSheetByName('CHECAGEM');
  var login = ss.getSheetByName('LOGIN');

  ValidacaoFiltros(base)
  ReexibirColunasOcultas(base)
  prepararBase(base)
  classificaBase(base)

  checagem.getRange('H6').activate()

  ValidacaoFiltros(checagem)
  ReexibirColunasOcultas(checagem) 
  limparChecagem(checagem)
  copiarDadosBase(base, checagem)
  cruzarDados(checagem, login) 
  ValidacaoEfectivo(checagem)

  ExcluirLinhasAbaixoUltimaLinha(checagem)

  checagem.getRange(7, 13, checagem.getLastRow(), checagem.getMaxColumns()).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)

  checagem.getRange(7, 1, checagem.getLastRow(), 7).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)

  OcutarColunas(checagem, 2, 5)

  checagem.getRange('H1').setFormula('=today()-1');

  checagem.getRange('H6').activate()

  if (base) {
    ss.deleteSheet(base);
    Logger.log('A planilha "BASE" foi excluída com sucesso.');
  } else {
    Logger.log('A planilha "BASE" não foi encontrada.');
  }

}

function TratamentoDeDados(){
  var ss = SpreadsheetApp.getActive();
  var checagem = ss.getSheetByName('CHECAGEM');

  checagem.getRange(7, 8, checagem.getLastRow() - 7, 1).splitTextToColumns(SpreadsheetApp.TextToColumnsDelimiter.SEMICOLON);
  checagem.getRange('H:M').trimWhitespace();

  ReexibirColunasOcultas(checagem)

  ValidacaoFiltros(checagem)

  ValidacoesErros(checagem)

  RemoverHashtags(checagem)

  OcutarColunas(checagem, 2, 4)

  ConfirmarEfectivo()

  ConfirmarMercadoPago()

  Cabecalho(checagem)

  AplicarFormatacaoCondicionalCompleta(checagem)

  var criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo(['0']).build();
  Filtrar(checagem, 6, criteria)

  checagem.getRange('L4').setFormula('=(COUNTIF(j:j;"")-1)');

  checagem.getRange('H1').setFormula('=TODAY()-1');
}

function Formatacao(){
  var ss = SpreadsheetApp.getActive();
  var checagem = ss.getSheetByName('CHECAGEM');

  checagem.getRange('H:H').setNumberFormat('dd/MM/yyyy')
  checagem.getRange('B:B').setNumberFormat('dd/MM/yyyy')

  checagem.getRange('I:I').setNumberFormat('#,##0.00')

  valoresWise(checagem)

  ValidacaoFiltros(checagem)

  var filter = checagem.getRange(6, 1, checagem.getLastRow(), checagem.getMaxColumns()).createFilter();

  // Classifica Data A a Z
  filter.sort(2, true);

  ReexibirColunasOcultas(checagem)

}
