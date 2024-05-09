// SpreadsheetApp.getUi().alert(rangePayment.getValues()[0][0]);
// sheet.getRange('D3').setBackground('white');

// function showform() {
//   var userform=HtmlService.createTemplateFromFile('form').evaluate().setTitle('Cadatsro de Aluno');
//   // SpreadsheetApp.getUi().showSidebar(userform);
//   SpreadsheetApp.getUi().showModelessDialog(userform, 'Cadastro de Aluno');
// }

// function GestaoAlunos() {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet();
//   var planilha = sheet.getSheetByName('Gestão de Alunos');
//   SpreadsheetApp.setActiveSheet(planilha);
// }

// Função para exibir uma mensagem de confirmação ou aviso
function confirmOperation(activeRow, sheetName, message) {
  // Obtém a planilha especificada
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  // Obtém a coluna da célula ativa
  let activeColumn = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getColumn();
  // Define a coluna para exibir a mensagem (uma coluna à direita do checkbox)
  let messageColumn = activeColumn + 1;
  // Obtém a célula na coluna definida e na mesma linha que a célula ativa
  let counterCell = sheet.getRange(activeRow, messageColumn);
  // Define o valor da célula para exibir a mensagem
  counterCell.setValue(message);
  // Aplica todas as alterações pendentes na planilha
  SpreadsheetApp.flush();
  // Pausa a execução do script pela quantidade de milissegundos especificada
  Utilities.sleep(1500);
  // Define o valor da célula para vazio, removendo a mensagem
  counterCell.setValue('');
}

// Função para efetuar a confirmação dos pagamentos
function registerPayment(activeRow, sheetName) {
  // Obtém a planilha especificada
  let paymentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro de Pagamentos');
  // Obtém a célula na coluna especificada na mesma linha que a célula ativa
  let counterCell = paymentSheet.getRange('H' + activeRow);
  // Obtém as células especificadas na mesma linha
  let rangeToCopy = paymentSheet.getRange(activeRow, 1, 1, 2);
  // Obtém a data atual
  let currentDate = new Date();
  // Obtém as células especificadas na mesma linha
  let paymentRange = paymentSheet.getRange(activeRow, 4, 1, 2);
  // Obtém os valores como uma matriz de 1D
  let valuesToCopy = rangeToCopy.getValues()[0];
  // Convertendo os valores para números antes de realizar a soma
  let counterValue = parseInt(counterCell.getValue());
  let paymentCounterValue = parseInt(paymentRange.getValues()[0][0]);
  // Convertendo o valor para número antes de realizar a checagem
  let paymentValue = parseInt(paymentRange.getValues()[0][1]);

  // Verifica se os valores não forem números ou forem menores ou iguais a zero
  if (isNaN(paymentCounterValue) || isNaN(paymentValue) || paymentCounterValue <= 0 || paymentValue <= 0) {
    // Se houver, exibe mensagem de aviso
    confirmOperation(activeRow, 'Registro de Pagamentos', '⚠️ Os valores dos pagamentos devem ser números maiores que zero');
    // Encerra a função, pois foi encontrado valores incorretos
    return;
  }

  // Soma o valor da célula com o valor especificado
  let summedValue = counterValue + paymentCounterValue;
  // Adiciona a data atual ao final da matriz
  valuesToCopy.push(currentDate);
  // Adiciona os valores especificados ao final da matriz
  valuesToCopy = valuesToCopy.concat(paymentRange.getValues()[0]);
  // Adiciona o valor somado ao final da matriz
  valuesToCopy.push(summedValue);

  // Obtém a planilha especificada
  let reportPaymentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relatório de Pagamentos');
  // Obtém a última linha na planilha especificada
  let lastRow = reportPaymentSheet.getLastRow();
  // Define a próxima linha para colar os dados
  let nextRow = lastRow + 1;
  // Cola os valores na próxima linha disponível na na planilha especificada
  reportPaymentSheet.getRange(nextRow, 1, 1, valuesToCopy.length).setValues([valuesToCopy]);

  // Obtém a planilha especificada
  let studentManagementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gestão de Alunos');
  // Define a variável para armazenar o número da linha correspondente
  let studentRow;
  // Obtém o nome na célula da coluna especificada na mesma linha
  let studentName = valuesToCopy[0];
  // Obtem a linha com nomes da planilha especificada
  let data = studentManagementSheet.getRange('A:A').getValues();

  // Procura a linha correspondente na planilha especificada
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == studentName) {
      studentRow = i + 1;
      break;
    }
  }

  // Define o valor nas células especificadas na linha correspondente
  studentManagementSheet.getRange(studentRow, 6).setValue(summedValue);
  counterCell.setValue(summedValue);
  // Zera os valores das células de pagamento
  paymentSheet.getRange(activeRow, 4, 1, 2).setValue(0);
  // Exibe mensagem de Confirmação
  confirmOperation(activeRow, sheetName, '✅ Confirmado');
}

// Função para contabilizar a presença do aluno
function registerPresence(activeRow) {
  // Obtém a planilha especificada
  let studentManagementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gestão de Alunos');
  // Obtém a célula na coluna especificada na mesma linha que a célula ativa
  let counterCell = studentManagementSheet.getRange('F' + activeRow);
  // Copia as 2 primeiras células na mesma linha
  let rangeToCopy = studentManagementSheet.getRange(activeRow, 1, 1, 2);
  // Obtém a data atual
  let currentDate = new Date();
  // Obtém o valor da célula e subtrai 1
  let valueToCopyMinusOne = counterCell.getValue() - 1;
  // Obtém os valores como uma matriz de 1D
  let valuesToCopy = rangeToCopy.getValues()[0];
  // Adiciona a data atual ao final da matriz
  valuesToCopy.push(currentDate);
  // Adiciona o valor decrementado ao final da matriz
  valuesToCopy.push(valueToCopyMinusOne);

  // Obtém a planilha especificada onde os dados serão colados
  let taughtLessonsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Aulas Ministradas');
  // Obtém a última linha na planilha especificada
  let lastRow = taughtLessonsSheet.getLastRow();
  // Define a próxima linha para colar os dados
  let nextRow = lastRow + 1;
  // Cola os valores na próxima linha disponível na planilha especificada
  taughtLessonsSheet.getRange(nextRow, 1, 1, valuesToCopy.length).setValues([valuesToCopy]);

  // Obtém a planilha especificada
  let paymentRegistrySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro de Pagamentos');
  let rowInPaymentRegistry;
  // Obtém o nome na célula da coluna especificada na mesma linha
  let studentName = valuesToCopy[0];
  // Procura a linha correspondente na planilha especificada
  let data = paymentRegistrySheet.getRange('A:A').getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == studentName) {
      rowInPaymentRegistry = i + 1; // Retorna o número da linha (começando do 1)
      break;
    }
  }

  // Define o valor na celula definida
  paymentRegistrySheet.getRange(rowInPaymentRegistry, 8).setValue(valueToCopyMinusOne);
  // Decrementa o valor da célula em -1
  counterCell.setValue(valueToCopyMinusOne);
}

// Função para cadastrar um novo aluno
function registerStudent(activeRow, sheetName){
  // Obtém a planilha especificada
  let studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  // Obtém os dados da planilha
  let data = studentSheet.getRange('C2:C3').getValues();
  // Obtém a data atual
  let currentDate = new Date();
  // Define o valor do checkbox como falso
  let checkbox = false;

  // Verificação de dados obrigatórios em branco
  if (data[0] == '' || data[1] == '') {
    // Se houver, exibe mensagem de aviso
    confirmOperation(activeRow, sheetName, '⚠️ Faltam dados obrigatórios');
    // Encerra a função, pois foi encontrado dados incorretos
    return;
  }

  // Verificação de nomes que não contém apenas letras ou não tem pelo menos duas palavras
  if (!/^[\p{L}\s]+$/u.test(data[0]) || data[0].toString().trim().split(' ').length < 2) { // corresponde a qualquer caractere de letra em qualquer idioma
    // Se houver, exibir mensagem de aviso
    confirmOperation(activeRow, sheetName, '⚠️ O primeiro valor deve conter apenas letras e pelo menos dois nomes');
    // Encerra a função, pois foi encontrado dados incorretos
    return;
  }

  // Verifica se o número de telefone não contém apenas números ou tem menos de 11 dígitos
  if (!/^\d{11,}$/.test(data[1])) {
    // Se houver, exibir mensagem de aviso
    confirmOperation(activeRow, sheetName, '⚠️ O segundo valor deve conter apenas números e ter no mínimo 11 dígitos');
    // Encerra a função, pois foi encontrado dados incorretos
    return;
  }

  // Informações adicionais
  let additionalLessonData = [currentDate, checkbox, '', 0];
  let additionalPaymentData = [currentDate, 0, 0, checkbox, '', 0];
  let optionalLessonQuantityValue = studentSheet.getRange('C4').getValue();
  let optionalLessonAmountValue = studentSheet.getRange('C5').getValue();

  // Verifica se os valores especificados não estão vazios e não são zero
  if ((optionalLessonQuantityValue != '' && optionalLessonQuantityValue != 0) || (optionalLessonAmountValue != '' && optionalLessonAmountValue != 0)) {
    // Verifica se os valores especificados são números maiores que zero
    if (!isNaN(optionalLessonQuantityValue) && optionalLessonQuantityValue > 0 && !isNaN(optionalLessonAmountValue) && optionalLessonAmountValue > 0) {
      // Define o valor de pagamento
      additionalLessonData[3] = optionalLessonQuantityValue;
      additionalPaymentData[5] = optionalLessonQuantityValue;

      // Obtém a planilha especificada
      let reportPaymentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relatório de Pagamentos');
      // Obtém a última linha na planilha especificada
      let lastRow = reportPaymentSheet.getLastRow();
      // Define a próxima linha para colar os dados
      let nextRow = lastRow + 1;

      // Cria uma matriz com os dados principais e adicionais de aulas
      let reportPaymentData = data.slice();
      reportPaymentData.push(currentDate);
      reportPaymentData.push(optionalLessonQuantityValue);
      reportPaymentData.push(optionalLessonAmountValue);
      reportPaymentData.push(optionalLessonQuantityValue);

      // Cola os valores na próxima linha disponível na planilha especificada
      reportPaymentSheet.getRange(nextRow, 1, 1, reportPaymentData.length).setValues([reportPaymentData]);
    } else {
      // Se os valores especificados não forem números maiores que zero, exibe mensagem de aviso
      confirmOperation(activeRow, sheetName, '⚠️ Os valores em C4 e C5 devem ser números maiores que zero');
      // Encerra a função, pois foi encontrado dados incorretos
      return;
    }
  }

  // Cria uma matriz com os dados principais e adicionais de aulas
  let lessonData = [data.concat(additionalLessonData)];
  // Cria uma matriz com os dados principais e adicionais para pagamentos
  let paymentData = [data.concat(additionalPaymentData)];

  // Obtém a planilha especificada
  let studentManagementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gestão de Alunos');
  // Obtém a última linha da planilha
  let lastStudentRow = studentManagementSheet.getLastRow();
  // Define a próxima linha para inserção de dados
  let targetStudentRow = lastStudentRow + 1;

  // Obtém a planilha especificada
  let paymentRegistrySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro de Pagamentos');
  // Obtém a última linha da planilha
  let lastPaymentRow = paymentRegistrySheet.getLastRow();
  // Define a próxima linha para inserção de dados
  let targetPaymentRow = lastPaymentRow + 1;

  // Obtém o nome a ser testado (primeiro nome dos dados inseridos), converte para minúsculas e remove espaços em branco do início e do final
  let testName = String(data[0]).trim().toLowerCase();

  // Verificação de Duplicatas Simples
  let currentNames = studentManagementSheet.getRange(1,1,lastStudentRow,1).getValues();
  // Iteração sobre a lista de nomes já cadastrados
  for (let i = 0; i < currentNames.length; i++){
    // Obtém o nome atual da lista de nomes já cadastrados, converte para minúsculas e remove espaços em branco do início e do final
    let currentName = String(currentNames[i]).trim().toLowerCase();
    //  Verifica se algum dos nomes é igual ao nome a ser testado
    if (currentName == testName) {
      // Se houver, exibir mensagem de aviso
      confirmOperation(activeRow, sheetName, '⚠️ Aluno já cadastrado');
      // Encerra a função, pois foi encontrado uma duplicata
      return;
    }
  }

  // Define a célula definida como um checkbox
  paymentRegistrySheet.getRange(targetPaymentRow, 6).insertCheckboxes();
  // Escrever dados finais na planilha especificada
  paymentRegistrySheet.getRange(targetPaymentRow, 1, 1, paymentData[0].length).setValues(paymentData);

  // Define a célula definida como um checkbox
  studentManagementSheet.getRange(targetStudentRow, 4).insertCheckboxes();
  // Escrever dados finais na planilha especificada
  studentManagementSheet.getRange(targetStudentRow, 1, 1, lessonData[0].length).setValues(lessonData);

  // Limpa os dados nas células especificadas
  studentSheet.getRange('C2:C5').clearContent();
  // Exibe mensagem de Confirmação
  confirmOperation(activeRow, sheetName, '✅ Aluno cadastrado com sucesso');
}

// Função acionada quando uma edição é feita na planilha
function onEdit() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  // Obtém a célula ativa na planilha
  let activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
  // Obtém a notação A1 da célula ativa
  let reference = activeCell.getA1Notation();
  // Obtém o nome da planilha que contém a célula ativa
  let sheetName = activeCell.getSheet().getName();
  // Obtém o valor da célula ativa
  let activeValue = activeCell.getValue();
  // Obtém o número da linha da célula ativa
  let activeRow = activeCell.getRow();
  // let activeColumn = activeCell.getColumn();

// if (sheetName == 'Registro de Alunos' || sheetName == 'Registro de Pagamentos'){
//   SpreadsheetApp.getUi().alert(reference+""+activeCell.getValue());
  // Verifica se a célula ativa é a célula D3 da planilha especificada e se o valor da célula é verdadeiro
  if (reference == 'D3' && sheetName == 'Registro de Alunos' && activeValue == true){
    // Chama a função para cadastrar um novo aluno
    registerStudent(activeRow, sheetName);
    // Define o valor da célula ativa como falso após o processamento
    activeCell.setValue(false);
  }
// }

  // Verifica se a célula ativa está em na coluna D da planilha especificada e se o valor da célula é verdadeiro
  if ((reference.startsWith('D')) && sheetName == 'Gestão de Alunos' && activeValue == true) {
    // Chama as funções para atualizar o contador e confirmar a operação
    registerPresence(activeRow);
    // Exibe mensagem de Confirmação
    confirmOperation(activeRow, sheetName, '✅ Confirmado');
    // Define o valor da célula ativa como falso após o processamento
    activeCell.setValue(false);
  }

  // Verifica se a célula ativa está em na coluna F da planilha especificada e se o valor da célula é verdadeiro
  if ((reference.startsWith('F')) && sheetName == 'Registro de Pagamentos' && activeValue == true) {
    // Chama as funções para atualizar o contador e confirmar a operação
    registerPayment(activeRow, sheetName);
    // Define o valor da célula ativa como falso após o processamento
    activeCell.setValue(false);
  }
}
