# Gerador e Envio de Relatórios de Carteiras por Gestores

## Descrição

Este projeto em Python foi desenvolvido para automatizar o processo de envio diário de relatórios de operações cadastradas no módulo de Cobrança. Após o cadastramento das operações no sistema, o código é executado para extrair as operações inseridas no dia atual. O relatório gerado é segmentado automaticamente por tipo de carteira, organizando as informações de acordo com as responsabilidades de cada gestor.

Em seguida, o relatório é enviado por e-mail para os gestores responsáveis e colaboradores interessados, garantindo uma comunicação eficiente e precisa sobre as operações registradas em suas respectivas carteiras.


## Problema de Negócio

O setor de cadastramento de indicações enfrenta diariamente uma carga de tarefas repetitivas e manuais, como o envio de relatórios com as indicações cadastradas no dia. Após o cadastramento dos casos indicados pelo Banco Santander, os colaboradores precisam identificar as operações inseridas, dividir manualmente esses registros por carteiras e enviar relatórios individuais para os gestores responsáveis por cada carteira, além de outros colaboradores envolvidos no fluxo. Esse processo consome tempo, é propenso a erros humanos e reduz a produtividade da equipe.

A automação desenvolvida busca eliminar esse gargalo, otimizando o fluxo de trabalho do setor de indicações, automatizando a geração, segmentação e envio dos relatórios diários, permitindo que os colaboradores foquem em atividades de maior valor agregado.


## Tecnologias e Ferramentas utilizadas
- **SQL**: Linguagem utilizadas dentro da biblioteca pyodbc para extrações de dados tabulares do banco de dados SQL Server - Ramaprod.
- **Python**: Todo código é feito em linguagem python, desde a biblioteca para extração dos dados, carregamento e transformação, divisão da base e envio dos relatórios por e-mail.
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
