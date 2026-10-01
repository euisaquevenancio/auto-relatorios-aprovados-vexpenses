# Automação Relatórios Aprovados Vexpenses 🤖✈️

Automação destinada para a captura de relatórios aprovados no Vexpenses das mantenedoras **ABEC**, **SOME**, **UBEE** e **UNBEC**, desenvolvida com Node.js. A coleta é feita diretamente na própria ferramenta, garantindo a extração automatizada das informações necessárias.

O processo inicia com o acesso ao Vexpenses e autenticação do usuário, seguido da coleta dos números e valores dos relatórios; nome do solicitante e também o tipo (reembolso, adiantamento ou prestação de contas). Registros já processados são desconsiderados com base no arquivo `relatoriosIgnorados.txt`. Em seguida, cada relatório é acessado individualmente para a extração de dados como data de vencimento, forma de pagamento e valor. A partir de campos como **Data de Aprovação Gestor**, **Data de Vencimento** e **Centro de Custos**.

Por fim, todos os relatóriso processados são consolidados em uma **planilha**, que é gerada e aberta automaticamente ao final da execução.

## Variáveis de ambiente

Para executar este projeto é necessário adicionar as seguintes variáveis no seu arquivo `.env`, referentes ao login no Vexpenses:

`EMAIL` e `SENHA`.

## Execução

Para executar a automação, é necessário instalar todos os arquivos presentes neste repositório. Após isso, abra o editor de código (ou similar) na pasta do projeto, certifique-se de ter o Node.js instalado e execute o seguinte comando no terminal para instalar as dependências:

```bash
    npm install
```

Após a instalação, execute o comando abaixo no terminal para iniciar a automação:

```bash
    npm run autoVexpenses.js
```

## Autores

- *[@euisaquevenancio](https://euisaquevenancio.github.io/portfolio/) - 25/09/2026*

