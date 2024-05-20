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
function confirmOperation(activeCell, sheetName, message, duration) {

  let activeRow = activeCell.getRow();
  // Obtém a planilha especificada
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  // Define a coluna para exibir a mensagem (uma coluna à direita do checkbox)
  let messageColumn = activeCell.getColumn() + 1;
  // Obtém a célula para exibir a mensagem
  let messageCell = sheet.getRange(activeRow, messageColumn);
  // Define o valor da célula para exibir a mensagem
  messageCell.setValue(message);
  // Aplica todas as alterações pendentes na planilha
  SpreadsheetApp.flush();


  if (duration !== 0) {
    // Pausa a execução do script pela quantidade de milissegundos especificada
    Utilities.sleep(duration);
    // Define o valor da célula para vazio, removendo a mensagem
    messageCell.setValue('');
  }
}


function getStudentRowByName(testName, sheet) {

  let currentNames = sheet.getRange('A:A').getValues();

  for (let i = 0; i < currentNames.length; i++) {

    let currentName = String(currentNames[i]).trim().toLowerCase();

    if (currentName == testName) {

      return i + 1;
    }
  }

  return -1;
}


function ensureRowExists(sheet, row) {

  if (row > sheet.getMaxRows()) {

    sheet.insertRowsAfter(sheet.getMaxRows(), row - sheet.getMaxRows());
  }
}


function registerStudentIfNotFound(testName, sheet, additionalData = []) {

  let lastRow = sheet.getLastRow();

  let newRow = lastRow + 1;

  let newStudentData = [testName, ...additionalData];

  ensureRowExists(sheet, newRow);

  sheet.getRange(newRow, 1, 1, newStudentData.length).setValues([newStudentData]);

  return newRow;
}

// Função para efetuar o registro de pagamentos
function registerPayment(activeCell, sheetName) {

  confirmOperation(activeCell, sheetName, '⏳ Em andamento...', 0);

  let lock = LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    let activeRow = activeCell.getRow();
    // Obtém a planilha especificada
    let paymentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro de Pagamentos');
    // Obtém a célula na coluna especificada na mesma linha que a célula ativa
    let counterCell = paymentSheet.getRange('H' + activeRow);

    let counterValue = parseInt(counterCell.getValue());


    if (isNaN(counterValue)) {

      confirmOperation(activeCell, sheetName, '⚠️ O valor do Saldo de Aulas não é um número válido', 3000);

      return;
    }


    // Obtém as células especificadas na mesma linha
    let rangeToCopy = paymentSheet.getRange(activeRow, 1, 1, 3);
    // Obtém a data atual
    let currentDate = new Date();
    // Obtém as células especificadas na mesma linha
    let paymentRange = paymentSheet.getRange(activeRow, 4, 1, 2);
    // Obtém os valores como uma matriz de 1D
    let valuesToCopy = rangeToCopy.getValues()[0].slice(0, 2);
    // Convertendo os valores para números antes de realizar a soma
    let paymentCounterValue = parseInt(paymentRange.getValues()[0][0]);
    // Convertendo o valor para número antes de realizar a checagem
    let paymentValue = parseInt(paymentRange.getValues()[0][1]);

    // Verifica se os valores não forem números ou forem menores ou iguais a zero
    if (isNaN(paymentCounterValue) || isNaN(paymentValue) || paymentCounterValue <= 0 || paymentValue <= 0) {
      // Se houver, exibe mensagem de aviso
      confirmOperation(activeCell, sheetName, '⚠️ Os valores dos pagamentos devem ser números maiores que zero', 3000);
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
    let studentManagementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gestão de Alunos');

    // let studentName = valuesToCopy[0];
    let testName = String(valuesToCopy[0]).trim().toLowerCase();

    let studentRow = getStudentRowByName(testName, studentManagementSheet);

    if (studentRow == -1) {
      studentRow = registerStudentIfNotFound(testName, studentManagementSheet, [valuesToCopy[1], rangeToCopy.getValues()[0][2], false, '', counterValue]);
      studentManagementSheet.getRange(studentRow, 4).insertCheckboxes();
    }

    // Obtém a planilha especificada
    let reportPaymentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relatório de Pagamentos');
    // Obtém a última linha na planilha especificada
    let lastRow = reportPaymentSheet.getLastRow();
    // Define a próxima linha para colar os dados
    let nextRow = lastRow + 1;

    ensureRowExists(reportPaymentSheet, nextRow);
    // Cola os valores na próxima linha disponível na na planilha especificada
    reportPaymentSheet.getRange(nextRow, 1, 1, valuesToCopy.length).setValues([valuesToCopy]);

    // Define o valor nas células especificadas na linha correspondente
    studentManagementSheet.getRange(studentRow, 6).setValue(summedValue);
    counterCell.setValue(summedValue);
    // Zera os valores das células de pagamento
    paymentSheet.getRange(activeRow, 4, 1, 2).setValue(0);
    // Exibe mensagem de Confirmação
    confirmOperation(activeCell, sheetName, '✅ Confirmado', 1500);
  } finally {

    lock.releaseLock();
  }
}

// Função para contabilizar a presença do aluno
function registerPresence(activeCell, sheetName) {

  confirmOperation(activeCell, sheetName, '⏳ Em andamento...', 0);

  let lock = LockService.getScriptLock();

  lock.waitLock(30000); // Espera até 30 segundos para adquirir a trava


  try {

    let activeRow = activeCell.getRow();
    // Obtém a planilha especificada
    let studentManagementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gestão de Alunos');
    // Obtém a célula na coluna especificada na mesma linha que a célula ativa
    let counterCell = studentManagementSheet.getRange('F' + activeRow);

    let counterValue = parseInt(counterCell.getValue());


    if (isNaN(counterValue)) {

      confirmOperation(activeCell, sheetName, '⚠️ O valor do Saldo de Aulas não é um número válido', 3000);

      return;
    }

    // Copia as 2 primeiras células na mesma linha
    let rangeToCopy = studentManagementSheet.getRange(activeRow, 1, 1, 3);
    // Obtém a data atual
    let currentDate = new Date();
    // Obtém o valor da célula e subtrai 1
    let valueToCopyMinusOne = counterValue - 1;
    // Obtém os valores como uma matriz de 1D
    let valuesToCopy = rangeToCopy.getValues()[0].slice(0, 2);
    // Adiciona a data atual ao final da matriz
    valuesToCopy.push(currentDate);
    // Adiciona o valor decrementado ao final da matriz
    valuesToCopy.push(valueToCopyMinusOne);

    // Obtém a planilha especificada
    let paymentRegistrySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro de Pagamentos');

    let testName = String(valuesToCopy[0]).trim().toLowerCase();

    let rowInPaymentRegistry = getStudentRowByName(testName, paymentRegistrySheet);


    if (rowInPaymentRegistry == -1) {

      rowInPaymentRegistry = registerStudentIfNotFound(testName, paymentRegistrySheet, [valuesToCopy[1], rangeToCopy.getValues()[0][2], 0, 0, false, '', counterValue]);

      paymentRegistrySheet.getRange(rowInPaymentRegistry, 6).insertCheckboxes();
    }

    // Obtém a planilha especificada onde os dados serão colados
    let taughtLessonsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Aulas Ministradas');
    // Obtém a última linha na planilha especificada
    let lastRow = taughtLessonsSheet.getLastRow();
    // Define a próxima linha para colar os dados
    let nextRow = lastRow + 1;

    ensureRowExists(taughtLessonsSheet, nextRow);
    // Cola os valores na próxima linha disponível na planilha especificada
    taughtLessonsSheet.getRange(nextRow, 1, 1, valuesToCopy.length).setValues([valuesToCopy]);

    // Define o valor na celula definida
    paymentRegistrySheet.getRange(rowInPaymentRegistry, 8).setValue(valueToCopyMinusOne);
    // Decrementa o valor da célula em -1
    counterCell.setValue(valueToCopyMinusOne);

    confirmOperation(activeCell, sheetName, '✅ Confirmado', 1500);
  } finally {

    lock.releaseLock();
  }
}

// Função para cadastrar um novo aluno
function registerStudent(activeCell, sheetName) {

  confirmOperation(activeCell, sheetName, '⏳ Em andamento...', 0);

  let lock = LockService.getScriptLock();

  lock.waitLock(30000); // Espera até 30 segundos para adquirir a trava


  try {
    // Obtém a planilha especificada
    let studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    // Obtém os dados da planilha
    let data = studentSheet.getRange('C2:C3').getValues();
    // Obtém a data atual
    let currentDate = new Date();

    // Verificação de dados obrigatórios em branco
    if (data[0] == '' || data[1] == '') {
      // Se houver, exibe mensagem de aviso
      confirmOperation(activeCell, sheetName, '⚠️ Faltam dados obrigatórios', 3000);
      // Encerra a função, pois foi encontrado dados incorretos
      return;
    }

    // Verificação de nomes que não contém apenas letras ou não tem pelo menos duas palavras
    if (!/^[\p{L}\s]+$/u.test(data[0]) || data[0].toString().trim().split(' ').length < 2) { // corresponde a qualquer caractere de letra em qualquer idioma
      // Se houver, exibir mensagem de aviso
      confirmOperation(activeCell, sheetName, '⚠️ O primeiro valor deve conter apenas letras e pelo menos dois nomes', 3000);
      // Encerra a função, pois foi encontrado dados incorretos
      return;
    }

    // Verifica se o número de telefone não contém apenas números ou tem menos de 11 dígitos
    if (!/^\d{11,15}$/.test(data[1])) {
      // Se houver, exibir mensagem de aviso
      confirmOperation(activeCell, sheetName, '⚠️ O número de telefone deve conter apenas números e respeitar limite de digitos', 3000);
      // Encerra a função, pois foi encontrado dados incorretos
      return;
    }

    // Informações adicionais
    let additionalLessonData = [currentDate, false, '', 0];
    let additionalPaymentData = [currentDate, 0, 0, false, '', 0];
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
        ensureRowExists(reportPaymentSheet, nextRow);

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
        confirmOperation(activeCell, sheetName, '⚠️ Os valores em C4 e C5 devem ser números maiores que zero', 3000);
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

    let testName = String(data[0]).trim().toLowerCase();


    if (getStudentRowByName(testName, studentManagementSheet) != -1 || getStudentRowByName(testName, paymentRegistrySheet) != -1) {

      confirmOperation(activeCell, sheetName, '⚠️ Aluno já cadastrado', 3000);

      return;
    }

    // Obtém a última linha da planilha
    let lastPaymentRow = paymentRegistrySheet.getLastRow();
    // Define a próxima linha para inserção de dados
    let targetPaymentRow = lastPaymentRow + 1;


    ensureRowExists(paymentRegistrySheet, targetPaymentRow);

    // Define a célula definida como um checkbox
    paymentRegistrySheet.getRange(targetPaymentRow, 6).insertCheckboxes();
    // Escrever dados finais na planilha especificada
    paymentRegistrySheet.getRange(targetPaymentRow, 1, 1, paymentData[0].length).setValues(paymentData);


    ensureRowExists(studentManagementSheet, targetStudentRow);

    // Define a célula definida como um checkbox
    studentManagementSheet.getRange(targetStudentRow, 4).insertCheckboxes();
    // Escrever dados finais na planilha especificada
    studentManagementSheet.getRange(targetStudentRow, 1, 1, lessonData[0].length).setValues(lessonData);

    // Limpa os dados nas células especificadas
    studentSheet.getRange('C2:C5').clearContent();
    // Exibe mensagem de Confirmação
    confirmOperation(activeCell, sheetName, '✅ Aluno cadastrado com sucesso', 1500);
  } finally {

    lock.releaseLock();
  }
}


// Função acionada quando uma edição é feita na planilha
function onEdit() {

  let lock = LockService.getScriptLock();

  lock.waitLock(30000); // Espera até 30 segundos para adquirir a trava


  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet();
    // Obtém a célula ativa na planilha
    let activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
    // Obtém a notação A1 da célula ativa
    let reference = activeCell.getA1Notation();
    // Obtém o nome da planilha que contém a célula ativa
    let sheetName = activeCell.getSheet().getName();
    // Obtém o valor da célula ativa
    let activeValue = activeCell.getValue();

    let cellD3 = sheet.getSheetByName('Registro de Alunos').getRange('D3');
    let cellD3value = cellD3.getValue();

    // Verifica se a célula ativa é a célula D3 da planilha especificada e se o valor da célula é verdadeiro
    if ((reference == 'D3' || reference == 'C2' || reference == 'C3' || reference == 'C4' || reference == 'C5') && sheetName == 'Registro de Alunos' && cellD3value == true) {
      // Chama a função para cadastrar um novo aluno
      registerStudent(cellD3, sheetName);
      // Define o valor da célula ativa como falso após o processamento
      cellD3.setValue(false);
    }

    // Verifica se a célula ativa está em na coluna D da planilha especificada e se o valor da célula é verdadeiro
    if ((reference.startsWith('D')) && sheetName == 'Gestão de Alunos' && activeValue == true) {
      // Chama as funções para atualizar o contador e confirmar a operação
      registerPresence(activeCell, sheetName);
      // Define o valor da célula ativa como falso após o processamento
      activeCell.setValue(false);
    }

    let paymentFColumnCell = sheet.getSheetByName('Registro de Pagamentos').getRange('F' + activeCell.getRow());
    let paymentCheckValue = paymentFColumnCell.getValue();

    // Verifica se a célula ativa está em na coluna F da planilha especificada e se o valor da célula é verdadeiro
    if ((reference.startsWith('F') || reference.startsWith('E') || reference.startsWith('D')) && sheetName == 'Registro de Pagamentos' && paymentCheckValue == true) {
      // Chama as funções para atualizar o contador e confirmar a operação
      registerPayment(paymentFColumnCell, sheetName);
      // Define o valor da célula ativa como falso após o processamento
      paymentFColumnCell.setValue(false);
    }
  } finally {
    lock.releaseLock();
  }
}
