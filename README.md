# Gerador e Envio de Relatórios de Carteiras por Gestores

## Descrição

Este projeto Python foi desenvolvido para automatizar o processo de extração de dados de carteiras de operações de um banco de dados SQL Server, gerar relatórios filtrados para cada gestor e enviá-los via e-mail. A solução usa integração com o Microsoft Outlook para enviar os e-mails e salvar relatórios personalizados em formato Excel para cada responsável pelas carteiras.

## Problema de Negócio

O objetivo deste código é auxiliar na distribuição automática de relatórios de acompanhamento de carteiras de casos gerenciadas por diferentes gestores dentro de uma empresa de advocacia. Ao invés de gerar manualmente relatórios diários e enviá-los para cada gestor, o código automatiza a geração e envio de relatórios detalhados, baseados em carteiras específicas, com dados atualizados diretamente do banco de dados.

### Aplicações
- **Geração automatizada de relatórios diários**: Gera relatórios personalizados com base nas carteiras atribuídas a cada gestor.
- **Envio de relatórios via e-mail**: Cada gestor recebe um e-mail com seu relatório de carteiras, incluindo um anexo Excel e uma tabela HTML com os dados no corpo do e-mail.
- **Automação de processos**: Reduz o esforço manual de extração e distribuição de dados, economizando tempo e evitando erros humanos.

## Lógica do Código

O código segue os seguintes passos:

1. **Carregamento de Variáveis de Ambiente**:
   - O arquivo `.env` é utilizado para armazenar credenciais de conexão com o banco de dados (servidor, usuário, senha e nome do banco).

2. **Conexão ao Banco de Dados SQL Server**:
   - O código usa `pyodbc` para conectar-se ao SQL Server e realizar uma consulta SQL. A consulta recupera dados de operações e carteiras, filtrando as operações do dia atual.

3. **Tratamento e Filtragem dos Dados**:
   - Os dados retornados são convertidos em um DataFrame do Pandas. Em seguida, o código filtra os dados de acordo com as carteiras gerenciadas por cada gestor, criando um arquivo Excel separado para cada um.

4. **Envio de E-mails Personalizados**:
   - O código usa a biblioteca `win32com` para se conectar ao Outlook, onde envia um e-mail para cada gestor com:
     - O arquivo Excel correspondente em anexo.
     - Uma tabela HTML gerada a partir dos dados no corpo do e-mail.
     - Uma assinatura com imagem personalizada.
  
## Pré-requisitos

Para rodar o código, você precisará:

- Python 3.x
- As seguintes bibliotecas Python:
  - `pandas`
  - `pyodbc`
  - `python-dotenv`
  - `win32com.client` (parte do pacote `pywin32`)
  
Além disso, você deve ter:
- Um banco de dados SQL Server com os dados de operações e carteiras.
- Microsoft Outlook instalado e configurado na máquina que executará o código.

## Como Usar

### 1. Clonar o Repositório
Clone o repositório do GitHub para sua máquina local:
```bash
git clone https://github.com/pedrocecere/nome-do-repositorio.git
