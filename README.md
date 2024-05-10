**README.md**

## Google Apps Script para Sistema de Gerenciamento de Estudantes

Este script foi desenvolvido para gerenciar dados de estudantes, presença em aulas e registros de pagamento dentro do Google Sheets. Ele fornece funções automatizadas para registrar estudantes, acompanhar sua presença e registrar pagamentos.

### Funções:

#### 1. `registerStudent(activeRow, sheetName)`

Esta função registra um novo estudante validando e armazenando suas informações na planilha especificada. Ela verifica dados obrigatórios, como nome e telefone, e opcionalmente permite especificar quantidades e valores adicionais de aulas.

#### 2. `registerPresence(activeRow)`

Registra a presença de um estudante em uma aula, decrementando o contador de aulas correspondente. Registra o nome do estudante, a data atual e a contagem de aulas decrementada em uma planilha separada.

#### 3. `registerPayment(activeRow, sheetName)`

Registra um pagamento feito por um estudante, atualizando os registros de pagamento e os saldos dos estudantes. Garante que os valores do pagamento sejam válidos e maiores que zero.

#### 4. `confirmOperation(activeRow, sheetName, message)`

Confirma uma operação exibindo uma mensagem na planilha especificada após a execução de uma ação. Fornece feedback visual ao usuário sobre o sucesso ou falha de uma operação.

#### 5. `onEdit()`

Esta é uma função de gatilho que é executada automaticamente quando edições são feitas na planilha. Detecta alterações em células específicas e aciona ações correspondentes, como registrar estudantes, registrar presença ou processar pagamentos.

### Como Usar:

1. **Configurando a Planilha:**
   - Crie um documento do Google Sheets e nomeie as planilhas conforme necessário: "Registro de Alunos", "Gestão de Alunos", "Aulas Ministradas", "Registro de Pagamentos" e "Relatório de Pagamentos".

2. **Adicionando o Script:**
   - Abra o documento do Google Sheets e vá para `Extensões` > `Apps Script`.
   - Copie e cole o script fornecido no editor do Apps Script.

3. **Executando Funções:**
   - As funções podem ser executadas manualmente ou acionadas automaticamente com base em edições específicas nas células.
   - Certifique-se de autorizar corretamente o script para acessar e modificar a planilha.

4. **Customização:**
   - Modifique o script conforme necessário para atender a requisitos específicos ou integrar funcionalidades adicionais.

### Observação:

- Este script foi projetado para funcionar com Google Sheets e Google Apps Script.
- Certifique-se de testar completamente a funcionalidade em um ambiente controlado antes de implantá-lo em um ambiente de produção.

Para informações detalhadas sobre cada função e como usá-las, consulte os comentários dentro do script.
