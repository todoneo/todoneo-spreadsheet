# Planilha de Gestão de Alunos e Financeira para Profissionais Autônomos

Esta planilha foi desenvolvida no Google Sheets para auxiliar profissionais autônomos, como professores, instrutores e consultores, a gerenciar seus alunos, acompanhar sua presença em aulas e pagamentos de forma simples e eficiente.

## Funcionalidades

* **Cadastro de Alunos:** Permite registrar novos alunos com informações básicas como nome completo, número de telefone e data de ingresso.
* **Registro de Pagamentos:** Registra os pagamentos dos alunos, incluindo a quantidade de aulas pagas, valor total e data do pagamento.
* **Gestão de Pagamentos:** Acompanha o histórico de pagamentos recebidos, incluindo a data do pagamento e o aluno correspondente.
* **Controle de Presença:** Permite marcar a presença dos alunos nas aulas, debitando automaticamente do saldo de aulas disponíveis.
* **Gestão de Aulas:** Acompanha o histórico de aulas ministradas, incluindo a data da aula e o aluno presente.
* **Dashboard Interativo:** Apresenta um resumo visual dos principais dados, como total de alunos, pagamentos recebidos, aulas ministradas e gráficos de desempenho.

## Estrutura da Planilha

A planilha é composta por várias abas:

* **Dashboard:** Apresenta um resumo visual dos principais dados.
* **Registro de Alunos:** Utilize esta aba para cadastrar novos alunos.
* **Gestão de Alunos:** Gerencie as informações dos alunos, como saldo de aulas e presença.
* **Registro de Pagamentos:** Registre os pagamentos recebidos dos alunos nesta aba.
* **Relatório de Pagamentos:** Visualize o histórico completo de pagamentos recebidos.
* **Aulas Ministradas:** Registre as aulas ministradas e a presença dos alunos.

## Como Utilizar

1. **Faça uma cópia da planilha:** Clique em "Arquivo" > "Fazer uma cópia" para criar uma cópia da planilha em sua conta do Google Drive.
2. **Comece a cadastrar seus alunos:** Acesse a aba "Registro de Alunos" e preencha os dados dos seus alunos.
3. **Registre os pagamentos:** Acesse a aba "Registro de Pagamentos" e registre os pagamentos recebidos dos seus alunos.
4. **Marque a presença dos alunos:** Na aba "Gestão de Alunos", marque a presença dos alunos nas aulas correspondentes.
5. **Acompanhe seu Dashboard:** A aba "Dashboard" será atualizada automaticamente com as informações cadastradas.

## Funções:

#### 1. `registerStudent(activeCell, sheetName)`

**Descrição:**

Esta função é responsável por registrar um novo aluno na planilha. Ela realiza validações nos dados de entrada, como a verificação de campos obrigatórios, formato do nome e número de telefone. Além disso, a função também permite registrar um pagamento inicial no momento do cadastro, caso haja. Os dados do novo aluno são adicionados às abas "Gestão de Alunos", "Registro de Pagamentos" e "Relatório de Pagamentos".

**Variáveis:**

* `activeCell`: A célula ativa na planilha, utilizada para exibir mensagens de status e erros.
* `sheetName`: O nome da aba da planilha onde a função foi chamada.

#### 2. `registerPresence(activeCell, sheetName)`

**Descrição:**

Esta função registra a presença de um aluno em uma aula. Ela verifica se o aluno existe na aba "Gestão de Alunos" e, em caso positivo, decrementa o saldo de aulas disponíveis. As informações da aula, como nome do aluno, data e saldo restante, são adicionadas à aba "Aulas Ministradas". Se o aluno não estiver cadastrado na aba "Registro de Pagamentos", ele será adicionado automaticamente.

**Variáveis:**

* `activeCell`: A célula ativa na planilha que acionou a função (checkbox de presença).
* `sheetName`: O nome da aba da planilha onde a função foi chamada.

#### 3. `registerPayment(activeCell, sheetName)`

**Descrição:**

Esta função registra um pagamento de um aluno. Ela valida os dados de entrada, garantindo que os valores sejam numéricos e maiores que zero. Em seguida, atualiza o saldo de aulas do aluno na aba "Gestão de Alunos" e adiciona o pagamento à aba "Relatório de Pagamentos". Se o aluno ainda não estiver cadastrado na aba "Gestão de Alunos", ele será adicionado automaticamente.

**Variáveis:**

* `activeCell`: A célula ativa na planilha que acionou a função (checkbox de pagamento).
* `sheetName`: O nome da aba da planilha onde a função foi chamada.

#### 4. `confirmOperation(activeCell, sheetName, message, duration)`

**Descrição:**

Esta função exibe uma mensagem de status ou erro em uma célula próxima à célula ativa. É utilizada para fornecer feedback visual ao usuário durante a execução das funções principais.

**Variáveis:**

* `activeCell`: A célula utilizada como referência para exibir a mensagem.
* `sheetName`: O nome da aba da planilha onde a mensagem será exibida.
* `message`: A mensagem a ser exibida.
* `duration`: O tempo em milissegundos que a mensagem será exibida. Se 0, a mensagem permanecerá visível até ser substituída por outra.

#### 5. `getStudentRowByName(testName, sheet)`

**Descrição:**

Esta função busca a linha de um aluno na planilha especificada usando o nome como critério de pesquisa.

**Variáveis:**

* `testName`: O nome do aluno a ser pesquisado.
* `sheet`: A aba da planilha onde a busca será realizada.

#### 6. `ensureRowExists(sheet, row)`

**Descrição:**

Esta função verifica se a linha especificada existe na planilha. Caso não exista, novas linhas são inseridas até que a linha desejada esteja disponível.

**Variáveis:**

* `sheet`: A aba da planilha onde a verificação será realizada.
* `row`: O número da linha a ser verificada.

#### 7. `registerStudentIfNotFound(testName, sheet, additionalData)`

**Descrição:**

Esta função registra um novo aluno na planilha caso ele ainda não esteja cadastrado.

**Variáveis:**

* `testName`: O nome do aluno a ser registrado.
* `sheet`: A aba da planilha onde o aluno será registrado.
* `additionalData`: Um array contendo informações adicionais do aluno.

#### 8. `updateAdmissionSummaryByMonth(startRow, startColumn)`

**Descrição:**

Esta função atualiza o resumo de ingressos de alunos por mês no Dashboard.

**Variáveis:**

* `startRow`: A linha onde o resumo deve começar a ser inserido.
* `startColumn`: A coluna onde o resumo deve começar a ser inserido.

#### 9. `updatePaymentSummaryByMonth(startRow, startColumn)`

**Descrição:**

Esta função atualiza o resumo de pagamentos recebidos por mês no Dashboard.

**Variáveis:**

* `startRow`: A linha onde o resumo deve começar a ser inserido.
* `startColumn`: A coluna onde o resumo deve começar a ser inserido.

#### 10. `updatePaymentSummary(startRow, startColumn)`

**Descrição:**

Esta função atualiza o resumo de pagamentos recebidos por aluno no Dashboard.

**Variáveis:**

* `startRow`: A linha onde o resumo deve começar a ser inserido.
* `startColumn`: A coluna onde o resumo deve começar a ser inserido.

#### 11. `updateLessonSummary(startRow, startColumn)`

**Descrição:**

Esta função atualiza o resumo de aulas ministradas por aluno no Dashboard.

**Variáveis:**

* `startRow`: A linha onde o resumo deve começar a ser inserido.
* `startColumn`: A coluna onde o resumo deve começar a ser inserido.

#### 12. `onEdit()`

**Descrição:**

Esta função é um gatilho que é executado automaticamente sempre que uma edição é feita na planilha. Ela verifica qual célula foi editada e executa as funções correspondentes, como registrar um novo aluno, registrar a presença em uma aula ou registrar um pagamento.

## Observações Importantes

* É necessário ter uma conta Google para utilizar a planilha.
* As funções da planilha são baseadas em scripts do Google Apps Script. Certifique-se de que os scripts estão habilitados em sua cópia da planilha.
* A planilha utiliza o recurso de trava de script para evitar conflitos de edição. Aguarde alguns segundos para que as operações sejam concluídas.

Para informações detalhadas sobre cada função e como usá-las, consulte os comentários dentro do script.
## Suporte e Sugestões

Para dúvidas, sugestões ou reportar erros, entre em contato com os desenvolvedores da Todoneo.

Site: [Todoneo](https://todoneo.com)
Contato: suporte@todoneo.com
Link da Planilha: [\[Planilha Todoneo\]](https://docs.google.com/spreadsheets/d/14fFBV8qHxBQtFm_fQJniuRVccC5F3TBTFBNRgPoPnpA/edit?usp=sharing)
