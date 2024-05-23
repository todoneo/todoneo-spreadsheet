// // Exibe um alerta
// SpreadsheetApp.getUi().alert(rangePayment.getValues()[0][0]);
// // Define o fundo da célula D3 como branco
// sheet.getRange('D3').setBackground('white');
// Exemplo de validação de dados mais robusta
// let dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
// let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Planilha');
// if (!dashboardSheet || !sheet) {
//   throw new Error("Planilha 'Dashboard' ou 'Aulas Ministradas' não encontrada.");
// }
// let dataRange = lessonsSheet.getRange('A4:A' + lessonsSheet.getLastRow());
// let data = dataRange.getValues();
// if (!data || data.length === 0) {
//   throw new Error("Nenhum dado encontrado na planilha 'Aulas Ministradas'.");
// }
// // Função para exibir um formulário
// function showform() {
//   // Cria e avalia um template HTML do arquivo 'form' e define o título da janela
//   var userform = HtmlService.createTemplateFromFile('form').evaluate().setTitle('Cadastro de Aluno');
//   // Exibe o formulário como uma janela modal
//   SpreadsheetApp.getUi().showModelessDialog(userform, 'Cadastro de Aluno');
// }
// // Função para encaminhar a visualização para uma planilha específica
// function GestaoAlunos() {
//   // Obtém a planilha ativa
//   var sheet = SpreadsheetApp.getActiveSpreadsheet();
//   // Obtém a planilha "Gestão de Alunos"
//   var planilha = sheet.getSheetByName('Gestão de Alunos');
//   // Define "Gestão de Alunos" como a planilha ativa
//   SpreadsheetApp.setActiveSheet(planilha);
// }
// // Funcional apenas com id nas tabelas
// function updateRegistroDePagamentos(activeCell, activeValue) {
//   let studentManagementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gestão de Alunos');
//   let paymentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro de Pagamentos');
//   let activeRow = activeCell.getRow();
//   let activeColumn = activeCell.getColumn();
//   let studentName = studentManagementSheet.getRange(activeRow, 1).getValue().trim().toLowerCase();
//   let paymentData = paymentSheet.getRange('A:H').getValues();
//   for (let i = 1; i < paymentData.length; i++) {
//     if (paymentData[i][0].trim().toLowerCase() == studentName) {
//       if (activeColumn == 1) { // Nome do Aluno
//         paymentSheet.getRange(i + 1, 1).setValue(activeValue);
//       } else if (activeColumn == 2) { // Telefone
//         paymentSheet.getRange(i + 1, 2).setValue(activeValue);
//       } else if (activeColumn == 6) { // Saldo de Aulas
//         paymentSheet.getRange(i + 1, 8).setValue(activeValue);
//       }
//       break;
//     }
//   }
// }
// Funcional apenas com id nas tabelas
// function updateGestaoDeAlunos(activeCell, activeValue) {
//   let studentManagementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gestão de Alunos');
//   let paymentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro de Pagamentos');
//   let activeRow = activeCell.getRow();
//   let activeColumn = activeCell.getColumn();
//   let studentName = paymentSheet.getRange(activeRow, 1).getValue().trim().toLowerCase();
//   let alunoData = studentManagementSheet.getRange('A:F').getValues();
//   for (let i = 1; i < alunoData.length; i++) {
//     if (alunoData[i][0].trim().toLowerCase() == studentName) {
//       if (activeColumn == 1) { // Nome do Aluno
//         studentManagementSheet.getRange(i + 1, 1).setValue(activeValue);
//       } else if (activeColumn == 2) { // Telefone
//         studentManagementSheet.getRange(i + 1, 2).setValue(activeValue);
//       } else if (activeColumn == 8) { // Saldo de Aulas
//         studentManagementSheet.getRange(i + 1, 6).setValue(activeValue);
//       }
//       break;
//     }
//   }
// }

// Função para exibir uma mensagem de confirmação ou aviso
function confirmOperation(activeCell, sheetName, message, duration) {
  // Obtém a planilha especificada
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  // Obtém a célula para exibir a mensagem (uma coluna à direita da célula especificada)
  let messageCell = sheet.getRange(activeCell.getRow(), activeCell.getColumn() + 1);
  // Define o valor da célula para exibir a mensagem
  messageCell.setValue(message);
  // Aplica todas as alterações pendentes na planilha
  SpreadsheetApp.flush();

  // Se a duração não for 0 executa funções específicas
  if (duration !== 0) {
    // Pausa a execução do script pela quantidade de milissegundos especificada
    Utilities.sleep(duration);
    // Limpa o conteúdo da célula após a pausa
    messageCell.clearContent();
  }
}

// Função para obter a linha de um aluno pelo nome
function getStudentRowByName(testName, sheet) {
  // Obtém uma lista com todos os nomes da coluna especificada
  let currentNames = sheet.getRange('A:A').getValues();

  // Itera pelos nomes na lista
  for (let i = 0; i < currentNames.length; i++) {
    // Obtém um nome da lista, remove espaços extras e converte para minúsculas
    let currentName = String(currentNames[i]).trim().toLowerCase();

    // Verifica se o nome atual é igual ao nome especificado
    if (currentName == testName) {
      // Se encontrar o nome, retorna a linha do nome
      return i + 1;
    }
  }

  // Se não encontrar o nome, retorna -1
  return -1;
}

// Função para garantir que uma linha epecificada existe na planilha
function ensureRowExists(sheet, row) {
  // Se a linha é maior que o número máximo de linhas da planilha
  if (row > sheet.getMaxRows()) {
    // Insere novas linhas até a linha especificada
    sheet.insertRowsAfter(sheet.getMaxRows(), row - sheet.getMaxRows());
  }
}

// Função para registrar um aluno se não encontrado
function registerStudentIfNotFound(testName, sheet, additionalData = []) {
  // Obtém a última linha com conteúdo da planilha especificada
  let lastRow = sheet.getLastRow();
  // Obtém a nova linha para o novo aluno
  let newRow = lastRow + 1;
  // Obtém uma matriz com o nome do aluno e dados adicionais
  let newStudentData = [testName, ...additionalData];
  // Garante que a linha especificada existe na planilha especificada
  ensureRowExists(sheet, newRow);
  // Insere os valores na nova linha
  sheet.getRange(newRow, 1, 1, newStudentData.length).setValues([newStudentData]);
  // Retorna a nova linha
  return newRow;
}

// Função para atualizar o resumo de ingresso de alunos por mês
function updateAdmissionSummaryByMonth(startRow, startColumn) {
  // Obtém uma trava de script
  let lock = LockService.getScriptLock();
  // Aguarda até 20 segundos para adquirir a trava
  lock.waitLock(20000);

  // Tenta executar o bloco de código
  try {
    // Obtém a planilha de alunos ativa e define o intervalo de dados
    let studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gestão de Alunos');
    // Obtém o intervalo de dados
    let dataRange = studentsSheet.getRange('C4:C' + studentsSheet.getLastRow());
    // Obtém os valores do intervalo
    let data = dataRange.getValues();
    // Cria um mapa para armazenar o resumo de ingressos por mês
    let admissionSummaryByMonth = {};

    // Itera sobre os dados
    data.forEach(row => {
      // Converte a data de admissão para um objeto Date
      let admissionDate = new Date(row[0]);

      // Verifica se a data é válida
      if (!isNaN(admissionDate.getTime())) {
        // Gera uma chave no formato 'mês/ano'
        let monthYearKey = (admissionDate.getMonth() + 1) + '/' + admissionDate.getFullYear();
        // Incrementa o contador de ingressos para o mês correspondente
        if (admissionSummaryByMonth[monthYearKey]) {
          // Se a chave já existir, incrementa o valor
          admissionSummaryByMonth[monthYearKey]++;
        } else {
          // Se a chave não existir, a inicializa com valor 1
          admissionSummaryByMonth[monthYearKey] = 1;
        }
      }
    });

    // Converte o mapa em uma matriz
    let summaryArray = Object.keys(admissionSummaryByMonth).map(monthYear => [monthYear, admissionSummaryByMonth[monthYear]]);
    // Ordena a matriz por data
    summaryArray.sort((a, b) => new Date(a[0].replace('/', '-')).getTime() - new Date(b[0].replace('/', '-')).getTime());

    // Obtém a planilha de dashboard e o intervalo alvo
    let dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
    // Obtém o intervalo alvo
    let targetRange = dashboardSheet.getRange(startRow, startColumn, summaryArray.length, 2);

    // Garante que haja linhas suficientes na planilha especificada
    ensureRowExists(dashboardSheet, startRow + summaryArray.length);
    // Insere os valores na planilha especificada
    targetRange.setValues(summaryArray);

    // Obtém o total de linhas da planilha especificada
    let maxRows = dashboardSheet.getMaxRows();
    // Verifica se há linhas abaixo da nova tabela
    if (startRow + summaryArray.length <= maxRows) {
      // Obtém o intervalo de células a serem limpas
      let clearRange = dashboardSheet.getRange(startRow + summaryArray.length, startColumn, maxRows - startRow - summaryArray.length + 1, 2);
      // Limpa o conteúdo das células
      clearRange.clearContent();
    }
  } finally {
    // Libera a trava de script
    lock.releaseLock();
  }
}

// Função para atualizar o resumo de pagamentos por mês
function updatePaymentSummaryByMonth(startRow, startColumn) {
  // Obtém uma trava de script
  let lock = LockService.getScriptLock();
  // Aguarda até 20 segundos para adquirir a trava
  lock.waitLock(20000);

  // Tenta executar o bloco de código
  try {
    // Obtém a planilha de relatório de pagamentos e define o intervalo de dados
    let paymentReportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relatório de Pagamentos');
    // Obtém o intervalo de dados
    let dataRange = paymentReportSheet.getRange('C4:E' + paymentReportSheet.getLastRow());
    // Obtém os valores do intervalo
    let data = dataRange.getValues();
    // Cria um mapa para armazenar o resumo de pagamentos por mês
    let paymentSummaryByMonth = {};

    // Itera sobre os dados
    data.forEach(row => {
      // Obtém a data do pagamento convertida para um objeto Date
      let paymentDate = new Date(row[0]);
      // Obtém o valor convertido para um número
      let paymentValue = parseFloat(row[2]);

      // Verifica se a data e o valor do pagamento são válidos
      if (!isNaN(paymentDate.getTime()) && !isNaN(paymentValue)) {
        // Obtém uma chave no formato 'mês/ano'
        let monthYearKey = (paymentDate.getMonth() + 1) + '/' + paymentDate.getFullYear();
        // Verifica se a chave já existe no mapa
        if (paymentSummaryByMonth[monthYearKey]) {
          // Se a chave já existir, incrementa o valor
          paymentSummaryByMonth[monthYearKey] += paymentValue;
        } else {
          // Se a chave não existir, a inicializa com o valor do pagamento
          paymentSummaryByMonth[monthYearKey] = paymentValue;
        }
      }
    });

    // Obtém uma matriz com os valores do mapa
    let summaryArray = Object.keys(paymentSummaryByMonth).map(monthYear => [monthYear, paymentSummaryByMonth[monthYear]]);
    // Ordena a matriz por data
    summaryArray.sort((a, b) => new Date(a[0].replace('/', '-')).getTime() - new Date(b[0].replace('/', '-')).getTime());

    // Obtém a planilha de dashboard e o intervalo alvo
    let dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
    // Obtém o intervalo alvo
    let targetRange = dashboardSheet.getRange(startRow, startColumn, summaryArray.length, 2);

    // Garante que haja linhas suficientes na planilha especificada
    ensureRowExists(dashboardSheet, startRow + summaryArray.length);
    // Insere os valores na planilha especificada
    targetRange.setValues(summaryArray);

    // Obtém o total de linhas da planilha especificada
    let maxRows = dashboardSheet.getMaxRows();
    // Verifica se há linhas abaixo da nova tabela
    if (startRow + summaryArray.length <= maxRows) {
      // Obtém o intervalo de células a serem limpas
      let clearRange = dashboardSheet.getRange(startRow + summaryArray.length, startColumn, maxRows - startRow - summaryArray.length + 1, 2);
      // Limpa o conteúdo das células
      clearRange.clearContent();
    }
  } finally {
    // Libera a trava de script
    lock.releaseLock();
  }
}

// Função para atualizar o resumo de pagamentos
function updatePaymentSummary(startRow, startColumn) {
  // Obtém uma trava de script
  let lock = LockService.getScriptLock();
  // Aguarda até 20 segundos para adquirir a trava
  lock.waitLock(20000);

  // Tenta executar o bloco de código
  try {
    // Obtém a planilha especificada
    let paymentReportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relatório de Pagamentos');
    // Obtém o intervalo de dados
    let dataRange = paymentReportSheet.getRange('A4:E' + paymentReportSheet.getLastRow());
    // Obtém os valores do intervalo
    let data = dataRange.getValues();

    // Cria um mapa para armazenar os totais pagos por aluno
    let paymentSummary = {};

    // Itera sobre os dados
    data.forEach(row => {
      // Obtém o nome do aluno
      let studentName = row[0];
      // Obtém o valor convertido para um número
      let paymentValue = parseFloat(row[4]);

      // Verifica se o nome do aluno e o valor do pagamento são válidos
      if (studentName && !isNaN(paymentValue)) {
        // Verifica se o nome do aluno já está no mapa
        if (paymentSummary[studentName]) {
          // Se existir, incrementa o valor
          paymentSummary[studentName] += paymentValue;
        } else {
          // Se a chave não existir, a inicializa com o valor do pagamento
          paymentSummary[studentName] = paymentValue;
        }
      }

    });

    // Obtém uma matriz com os valores do mapa
    let summaryArray = Object.keys(paymentSummary).map(student => [student, paymentSummary[student]]);
    // Ordena a matriz pelo nome do aluno
    summaryArray.sort((a, b) => a[0].localeCompare(b[0]));

    // Obtém a planilha especificada
    let dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
    // Obtém o intervalo alvo
    let targetRange = dashboardSheet.getRange(startRow, startColumn, summaryArray.length, 2);

    // Garante que haja linhas suficientes na planilha especificada
    ensureRowExists(dashboardSheet, startRow + summaryArray.length);
    // Insere os valores na planilha especificada
    targetRange.setValues(summaryArray);

    // Obtém o total de linhas da planilha especificada
    let maxRows = dashboardSheet.getMaxRows();
    // Verifica se há linhas abaixo da nova tabela
    if (startRow + summaryArray.length <= maxRows) {
      // Obtém o intervalo de células a serem limpas
      let clearRange = dashboardSheet.getRange(startRow + summaryArray.length, startColumn, maxRows - startRow - summaryArray.length + 1, 2);
      // Limpa o conteúdo das células
      clearRange.clearContent();
    }
  } finally {
    // Libera a trava de script
    lock.releaseLock();
  }
}

// Função para atualizar o resumo de aulas por aluno
function updateLessonSummary(startRow, startColumn) {
  // Obtém uma trava de script
  let lock = LockService.getScriptLock();
  // Aguarda até 20 segundos para adquirir a trava
  lock.waitLock(20000);

  // Tenta executar o bloco de código
  try {
    // Obtém a planilha especificada
    let lessonsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Aulas Ministradas');

    // Obtém o intervalo de dados
    let dataRange = lessonsSheet.getRange('A4:A' + lessonsSheet.getLastRow());
    // Obtém os valores do intervalo
    let data = dataRange.getValues();

    // Cria um mapa para armazenar o total de aulas por aluno
    let lessonSummary = {};

    data.forEach(row => {
      // Obtém o nome do aluno
      let studentName = row[0];

      // Verifica se o nome do aluno é válido
      if (studentName) {
        // Verifica se o nome do aluno já está no mapa
        if (lessonSummary[studentName]) {
          // Se existir, incrementa o valor
          lessonSummary[studentName] += 1;
        } else {
          // Se a chave não existir, a inicializa com o valor 1
          lessonSummary[studentName] = 1;
        }
      }
    });

    // Obtém uma matriz com os valores do mapa
    let summaryArray = Object.keys(lessonSummary).map(student => [student, lessonSummary[student]]);
    // Ordena a matriz pelo nome do aluno
    summaryArray.sort((a, b) => a[0].localeCompare(b[0]));

    // Obtém a planilha especificada
    let dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
    // Obtém o intervalo alvo
    let targetRange = dashboardSheet.getRange(startRow, startColumn, summaryArray.length, 2);

    // Garante que haja linhas suficientes na planilha especificada
    ensureRowExists(dashboardSheet, startRow + summaryArray.length);
    // Insere os valores na planilha especificada
    targetRange.setValues(summaryArray);

    // Obtém o total de linhas da planilha especificada
    let maxRows = dashboardSheet.getMaxRows();
    // Verifica se há linhas abaixo da nova tabela
    if (startRow + summaryArray.length <= maxRows) {
      // Obtém o intervalo de células a serem limpas
      let clearRange = dashboardSheet.getRange(startRow + summaryArray.length, startColumn, maxRows - startRow - summaryArray.length + 1, 2);
      // Limpa o conteúdo das células
      clearRange.clearContent();
    }
  } finally {
    // Libera a trava de script
    lock.releaseLock();
  }
}

// Função para efetuar o registro de pagamentos
function registerPayment(activeCell, sheetName) {
  // Obtém uma trava de script
  let lock = LockService.getScriptLock();
  // Aguarda até 20 segundos para adquirir a trava
  lock.waitLock(20000);

  // Tenta executar o bloco de código
  try {
    // Obtém a linha da célula ativa
    let activeRow = activeCell.getRow();
    // Obtém a planilha especificada
    let paymentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro de Pagamentos');
    // Obtém a célula na coluna especificada na mesma linha que a célula ativa
    let counterCell = paymentSheet.getRange('H' + activeRow);
    // Obtém o valor da célula e converte para inteiro
    let counterValue = parseInt(counterCell.getValue());

    // Verifica se o valor não é um número
    if (isNaN(counterValue)) {
      // Se houver, exibe uma mensagem de aviso
      confirmOperation(activeCell, sheetName, '⚠️ O valor do Saldo de Aulas não é um número válido', 3000);
      // Encerra a função, pois foi encontrado valores incorretos
      return;
    }


    // Obtém as células especificadas na linha ativa
    let rangeToCopy = paymentSheet.getRange(activeRow, 1, 1, 3);
    // Obtém a data atual
    let currentDate = new Date();
    // Obtém as células especificadas na linha ativa
    let paymentRange = paymentSheet.getRange(activeRow, 4, 1, 2);
    // Obtém os valores como uma matriz de 1D
    let valuesToCopy = rangeToCopy.getValues()[0].slice(0, 2);
    // Obtém os valores convertidos para números
    let paymentCounterValue = parseInt(paymentRange.getValues()[0][0]);
    let paymentValue = parseFloat(paymentRange.getValues()[0][1]);

    // Verifica se os valores não forem números ou forem menores ou iguais a zero
    if (isNaN(paymentCounterValue) || isNaN(paymentValue) || paymentCounterValue <= 0 || paymentValue <= 0) {
      // Se houver, exibe uma mensagem de aviso
      confirmOperation(activeCell, sheetName, '⚠️ Os valores dos pagamentos devem ser números maiores que zero', 3000);
      // Encerra a função, pois foi encontrado valores incorretos
      return;
    }

    // Obtém a soma do saldo de aulas e do valor especificado
    let summedValue = counterValue + paymentCounterValue;
    // Adiciona a data atual ao final da matriz
    valuesToCopy.push(currentDate);
    // Adiciona os valores especificados ao final da matriz
    valuesToCopy = valuesToCopy.concat(paymentRange.getValues()[0]);
    // Adiciona o valor somado ao final da matriz
    valuesToCopy.push(summedValue);

    // Obtém a planilha especificada
    let studentManagementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gestão de Alunos');

    // Obtém o nome do aluno convertido para caracteres minúsculas
    let testName = String(valuesToCopy[0]).trim().toLowerCase();
    // Obtém a linha do aluno pelo nome na planilha especificada
    let studentRow = getStudentRowByName(testName, studentManagementSheet);

    // Verifica se o aluno não está cadastrado na planilha especificada
    if (studentRow == -1) {
      // Se não encontrar o aluno, registra o aluno
      studentRow = registerStudentIfNotFound(testName, studentManagementSheet, [valuesToCopy[1], rangeToCopy.getValues()[0][2], false, '', counterValue]);
      // Define a célula especificada como um checkbox
      studentManagementSheet.getRange(studentRow, 4).insertCheckboxes();
    }

    // Obtém a planilha especificada
    let reportPaymentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relatório de Pagamentos');
    // Obtém a última linha com conteúdo da planilha especificada
    let lastRow = reportPaymentSheet.getLastRow();
    // Garante que a linha especificada existe na planilha especificada
    ensureRowExists(reportPaymentSheet, lastRow + 2);
    // Insere os valores na próxima linha disponível na planilha especificada
    reportPaymentSheet.getRange(lastRow + 1, 1, 1, valuesToCopy.length).setValues([valuesToCopy]);

    // Define o valor nas células especificadas
    studentManagementSheet.getRange(studentRow, 6).setValue(summedValue);
    counterCell.setValue(summedValue);
    // Zera os valores das células de pagamento
    paymentSheet.getRange(activeRow, 4, 1, 2).setValue(0);

    // Exibe mensagem de confirmação
    confirmOperation(activeCell, sheetName, '✅ Pagamento registrado com sucesso', 1500);
  } finally {
    // Libera a trava de script
    lock.releaseLock();
  }
}

// Função para contabilizar a presença do aluno
function registerPresence(activeCell, sheetName) {
  // Obtém uma trava de script
  let lock = LockService.getScriptLock();
  // Aguarda até 20 segundos para adquirir a trava
  lock.waitLock(20000);

  // Tenta executar o bloco de código
  try {
    // Obtém a linha da célula ativa
    let activeRow = activeCell.getRow();
    // Obtém a planilha especificada
    let studentManagementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gestão de Alunos');
    // Obtém a célula na coluna especificada na mesma linha que a célula ativa
    let counterCell = studentManagementSheet.getRange('F' + activeRow);
    // Obtém o valor da célula convertido para inteiro
    let counterValue = parseInt(counterCell.getValue());

    // Verifica se o valor não é um número
    if (isNaN(counterValue)) {
      // Se houver, exibe mensagem de aviso
      confirmOperation(activeCell, sheetName, '⚠️ O valor do Saldo de Aulas não é um número válido', 3000);
      // Encerra a função, pois foi encontrado valores incorretos
      return;
    }

    // Obtém as células especificadas na linha ativa
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
    // Obtém o nome do aluno convertido para caracteres minúsculas
    let testName = String(valuesToCopy[0]).trim().toLowerCase();
    // Obtém a linha do aluno pelo nome na planilha especificada
    let rowInPaymentRegistry = getStudentRowByName(testName, paymentRegistrySheet);

    // Verifica se o aluno não está cadastrado na planilha especificada
    if (rowInPaymentRegistry == -1) {
      // Se não encontrar o aluno, registra o aluno
      rowInPaymentRegistry = registerStudentIfNotFound(testName, paymentRegistrySheet, [valuesToCopy[1], rangeToCopy.getValues()[0][2], 0, 0, false, '', counterValue]);
      // Define a célula especificada como um checkbox
      paymentRegistrySheet.getRange(rowInPaymentRegistry, 6).insertCheckboxes();
    }

    // Obtém a planilha especificada onde os dados serão colados
    let taughtLessonsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Aulas Ministradas');
    // Obtém a última linha com conteúdo da planilha especificada
    let lastRow = taughtLessonsSheet.getLastRow();
    // Garante que a linha especificada existe na planilha especificada
    ensureRowExists(taughtLessonsSheet, lastRow + 2);
    // Insere os valores na próxima linha disponível na planilha especificada
    taughtLessonsSheet.getRange(lastRow + 1, 1, 1, valuesToCopy.length).setValues([valuesToCopy]);

    // Define o valor na celula especificada
    paymentRegistrySheet.getRange(rowInPaymentRegistry, 8).setValue(valueToCopyMinusOne);
    // Decrementa o valor da célula em -1
    counterCell.setValue(valueToCopyMinusOne);

    // Exibe mensagem de confirmação
    confirmOperation(activeCell, sheetName, '✅ Presença registrada com sucesso', 1500);
  } finally {
    // Libera a trava de script
    lock.releaseLock();
  }
}

// Função para cadastrar um novo aluno
function registerStudent(activeCell, sheetName) {
  // Obtém uma trava de script
  let lock = LockService.getScriptLock();
  // Aguarda até 20 segundos para adquirir a trava
  lock.waitLock(20000);

  // Tenta executar o bloco de código
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
      confirmOperation(activeCell, sheetName, '⚠️ Nome completo do aluno ou número de telefone estão em branco', 3000);
      // Encerra a função, pois foi encontrado dados incorretos
      return;
    }

    // Verificação de nomes que não contém apenas letras ou não tem pelo menos duas palavras, corresponde a qualquer caractere de letra em qualquer idioma
    if (!/^[\p{L}\s]+$/u.test(data[0]) || data[0].toString().trim().split(' ').length < 2) {
      // Se houver, exibir mensagem de aviso
      confirmOperation(activeCell, sheetName, '⚠️ O primeiro valor deve conter apenas letras e pelo menos dois nomes', 3000);
      // Encerra a função, pois foi encontrado dados incorretos
      return;
    }

    // Verifica se o número de telefone não contém apenas números ou tem menos de 11 dígitos
    if (!/^\d{11,15}$/.test(data[1])) {
      // Se houver, exibir mensagem de aviso
      confirmOperation(activeCell, sheetName, '⚠️ O número de telefone deve conter apenas números e respeitar limite de dígitos', 3000);
      // Encerra a função, pois foi encontrado dados incorretos
      return;
    }

    // Obtém informações adicionais
    let additionalLessonData = [currentDate, false, '', 0];
    let additionalPaymentData = [currentDate, 0, 0, false, '', 0];
    let optionalLessonQuantityValue = studentSheet.getRange('C4').getValue();
    let optionalLessonAmountValue = studentSheet.getRange('C5').getValue();

    // Verifica se os valores especificados não estão vazios e não são zero
    if ((optionalLessonQuantityValue != '' && optionalLessonQuantityValue != 0) || (optionalLessonAmountValue != '' && optionalLessonAmountValue != 0)) {
      // Verifica se os valores especificados são números maiores que zero
      if (!isNaN(optionalLessonQuantityValue) && optionalLessonQuantityValue > 0 && !isNaN(optionalLessonAmountValue) && optionalLessonAmountValue > 0) {
        // Adiciona os valores à matriz de dados adicionais
        additionalLessonData[3] = optionalLessonQuantityValue;
        additionalPaymentData[5] = optionalLessonQuantityValue;

        // Obtém a planilha especificada
        let reportPaymentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relatório de Pagamentos');
        // Obtém a última linha com conteúdo da planilha especificada
        let lastRow = reportPaymentSheet.getLastRow();
        // Garante que a linha especificada existe na planilha especificada
        ensureRowExists(reportPaymentSheet, lastRow + 2);

        // Obtém uma matriz com os dados principais e adicionais de aulas
        let reportPaymentData = data.slice();
        reportPaymentData.push(currentDate);
        reportPaymentData.push(optionalLessonQuantityValue);
        reportPaymentData.push(optionalLessonAmountValue);
        reportPaymentData.push(optionalLessonQuantityValue);

        // Insere os valores na próxima linha disponível na planilha especificada
        reportPaymentSheet.getRange(lastRow + 1, 1, 1, reportPaymentData.length).setValues([reportPaymentData]);
      } else {
        // Se os valores especificados não forem números maiores que zero, exibe mensagem de aviso
        confirmOperation(activeCell, sheetName, '⚠️ Os valores em C4 e C5 devem ser números maiores que zero', 3000);
        // Encerra a função, pois foi encontrado dados incorretos
        return;
      }
    }

    // Obtém uma matriz com os dados principais e adicionais de aulas
    let lessonData = [data.concat(additionalLessonData)];
    // Obtém uma matriz com os dados principais e adicionais para pagamentos
    let paymentData = [data.concat(additionalPaymentData)];

    // Obtém a planilha especificada
    let studentManagementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gestão de Alunos');
    // Obtém a última linha com conteúdo da planilha especificada
    let lastStudentRow = studentManagementSheet.getLastRow();
    // Obtém a próxima linha para inserção de dados
    let targetStudentRow = lastStudentRow + 1;

    // Obtém a planilha especificada
    let paymentRegistrySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro de Pagamentos');
    // Obtém o nome do aluno convertido para caracteres minúsculas
    let testName = String(data[0]).trim().toLowerCase();

    // Verifica se o aluno já está cadastrado
    if (getStudentRowByName(testName, studentManagementSheet) != -1 || getStudentRowByName(testName, paymentRegistrySheet) != -1) {
      // Se houver, exibe mensagem de aviso
      confirmOperation(activeCell, sheetName, '⚠️ Aluno já cadastrado', 3000);
      // Encerra a função, pois foi encontrado duplicata
      return;
    }

    // Obtém a última linha com conteúdo da planilha especificada
    let lastPaymentRow = paymentRegistrySheet.getLastRow();
    // Obtém a próxima linha para inserção de dados
    let targetPaymentRow = lastPaymentRow + 1;

    // Garante que a linha especificada existe na planilha especificada
    ensureRowExists(paymentRegistrySheet, targetPaymentRow);

    // Define a célula especificada como um checkbox
    paymentRegistrySheet.getRange(targetPaymentRow, 6).insertCheckboxes();
    // Insere os dados na planilha especificada
    paymentRegistrySheet.getRange(targetPaymentRow, 1, 1, paymentData[0].length).setValues(paymentData);

    // Garante que a linha especificada existe na planilha especificada
    ensureRowExists(studentManagementSheet, targetStudentRow);

    // Define a célula especificada como um checkbox
    studentManagementSheet.getRange(targetStudentRow, 4).insertCheckboxes();
    // Insere os dados na planilha especificada
    studentManagementSheet.getRange(targetStudentRow, 1, 1, lessonData[0].length).setValues(lessonData);

    // Limpa os dados nas células especificadas
    studentSheet.getRange('C2:C5').clearContent();
    // Exibe mensagem de confirmação
    confirmOperation(activeCell, sheetName, '✅ Aluno registrado com sucesso', 1500);
  } finally {
    // Libera a trava de script
    lock.releaseLock();
  }
}


// Função acionada quando uma edição é feita na planilha
function onEdit() {
  // Obtém uma trava de script
  let lock = LockService.getScriptLock();
  // Aguarda até 20 segundos para adquirir a trava
  lock.waitLock(20000);

  // Tenta executar o bloco de código
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet();
    // Obtém a célula ativa na planilha
    let activeCell = sheet.getActiveCell();
    // Obtém a notação A1 da célula ativa
    let reference = activeCell.getA1Notation();
    // Obtém o nome da planilha que contém a célula ativa
    let sheetName = activeCell.getSheet().getName();
    // Obtém o valor da célula ativa
    let activeValue = activeCell.getValue();

    // Obtém a célula especificada
    let cellD3 = sheet.getSheetByName('Registro de Alunos').getRange('D3');
    // Obtém o valor da célula especificada
    let cellD3value = cellD3.getValue();

    // Verifica se a célula ativa é a célula D3 da planilha especificada e se o valor da célula é verdadeiro
    if ((reference == 'D3' || reference == 'C2' || reference == 'C3' || reference == 'C4' || reference == 'C5') && sheetName == 'Registro de Alunos' && cellD3value == true) {
      // Tenta executar o bloco de código
      try {
        // Exibe mensagem de execução
        confirmOperation(cellD3, sheetName, '⏳ Em andamento...', 0); // '⏳ Registrando novo aluno. Aguarde...'
        // Chama a função para cadastrar um novo aluno
        registerStudent(cellD3, sheetName);
      } catch (error) {
        // Exibe mensagem de erro
        confirmOperation(cellD3, sheetName, '⚠️ Erro ao registrar aluno, cheque as informações do aluno e tente novamente', 3000);
      } finally {
        // Define o valor da célula como falso
        cellD3.setValue(false);
        // Limpa qualquer mensagem no campo
        confirmOperation(cellD3, sheetName, '', 0);
      }
    }

    // Verifica se a célula ativa está em na coluna D da planilha especificada e se o valor da célula é verdadeiro
    if ((reference.startsWith('D')) && sheetName == 'Gestão de Alunos' && activeValue == true) {
      // Tenta executar o bloco de código
      try {
        // Exibe mensagem de execução
        confirmOperation(activeCell, sheetName, '⏳ Em andamento...', 0); // '⏳ Registrando presença do aluno...'
        // Chama a função para registrar presença do aluno na aula
        registerPresence(activeCell, sheetName);
        // Define o valor da célula ativa como falso após o processamento
        activeCell.setValue(false);
      } catch (error) {
        // Exibe mensagem de erro
        confirmOperation(activeCell, sheetName, '⚠️ Erro ao registrar presença, cheque as informações do aluno e tente novamente', 3000);
      } finally {
        // Define o valor da célula como falso
        activeCell.setValue(false);
        // Limpa qualquer mensagem no campo
        confirmOperation(activeCell, sheetName, '', 0);
      }
    }

    // Obtém a célula especificada
    let paymentFColumnCell = sheet.getSheetByName('Registro de Pagamentos').getRange('F' + activeCell.getRow());
    // Obtém o valor da célula especificada
    let paymentCheckValue = paymentFColumnCell.getValue();

    // Verifica se a célula ativa está em na coluna F da planilha especificada e se o valor da célula é verdadeiro
    if ((reference.startsWith('F') || reference.startsWith('E') || reference.startsWith('D')) && sheetName == 'Registro de Pagamentos' && paymentCheckValue == true) {
      // Tenta executar o bloco de código
      try {
        // Exibe mensagem de execução
        confirmOperation(paymentFColumnCell, sheetName, '⏳ Em andamento...', 0); // '⏳ Registrando pagamento...'
        // Chama a função para registrar pagamento
        registerPayment(paymentFColumnCell, sheetName);
      } catch (error) {
        // Exibe mensagem de erro
        confirmOperation(paymentFColumnCell, sheetName, '⚠️ Erro ao registrar pagamento, cheque as informações do aluno e tente novamente', 3000);
      } finally {
        // Define o valor da célula como falso
        paymentFColumnCell.setValue(false);
        // Limpa qualquer mensagem no campo
        confirmOperation(paymentFColumnCell, sheetName, '', 0);
      }
    }

    // Verifica se a célula ativa é a célula A2 da planilha especificada e se o valor da célula é verdadeiro
    if (reference == 'A2' && sheetName == 'Dashboard' && activeValue == true) {
      // Tenta executar o bloco de código
      try {
        // Exibe mensagem de execução
        confirmOperation(activeCell, sheetName, '⏳ Em andamento...', 0); // '⏳ Atualizando Dashboard...'
        // Chama a função para atualizar a tabela de total pago por aluno
        updatePaymentSummary(4, 11);
        // Chama a função para atualizar a tabela de total de aulas por aluno
        updateLessonSummary(4, 14);
        // Chama a função para atualizar a tabela de total de aulas por mês
        updatePaymentSummaryByMonth(4, 17);
        // Chamada a função para atualizar a tabela de total de ingressos por mês
        updateAdmissionSummaryByMonth(4, 20);

        // Exibe mensagem de confirmação
        confirmOperation(activeCell, sheetName, '✅ Dashboard atualizado', 1500);
      } catch (error) {
        // Exibe mensagem de erro
        confirmOperation(activeCell, sheetName, '⚠️ Erro ao atualizar Dashboard, cheque as informações do aluno e tente novamente', 3000);
      } finally {
        // Define o valor da célula como falso
        activeCell.setValue(false);
        // Limpa qualquer mensagem no campo
        confirmOperation(activeCell, sheetName, '', 0);
      }
    }
    // Funcional apenas com id nas tabelas
    // if ((reference.startsWith('A') || reference.startsWith('B') || reference.startsWith('F')) && sheetName == 'Gestão de Alunos') {
    //   updateRegistroDePagamentos(activeCell, activeValue);
    // } else if ((reference.startsWith('A') || reference.startsWith('B') || reference.startsWith('H')) && sheetName == 'Registro de Pagamentos') {
    //   updateGestaoDeAlunos(activeCell, activeValue);
    // }
  } finally {
    // Libera a trava de script
    lock.releaseLock();
  }
}
