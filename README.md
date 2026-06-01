# Sistema de Gestao de Copias

Dashboard em Google Apps Script para acompanhar liberacoes de copias, relatorios operacionais, KPI de fechamento, retorno de requisitos, testes de copias e desempenho de operadores.

## Estrutura

- `codigo.gs`: back-end Apps Script, leitura/escrita no Google Sheets, API `doPost` e funcoes chamadas pelo front-end.
- `Index.html`: template principal do Web App.
- `Styles.html`: estilos do dashboard.
- `Scripts.html`: comportamento do front-end, graficos, filtros, KPI e exportacao.
- `appsscript.json`: manifest do Apps Script.

## Abas Esperadas

O projeto usa uma planilha Google Sheets com estas abas:

- `Geral`: base principal de copias.
- `KPI_Historico`: historico dos KPIs importados.
- `KPI_Sessoes`: auditoria das importacoes de KPI.
- `REQ_Historico`: historico das importacoes de retorno de requisitos.
- `REQ_Sessoes`: auditoria das importacoes de retorno de requisitos.
- `TESTES_Historico`: historico das importacoes de testes de copias.
- `TESTES_Sessoes`: auditoria das importacoes de testes de copias.
- `Operadores`: cadastro substituivel de operadores e nomes.
- `SYS_LOG`: log tecnico do sistema.

Se as abas auxiliares nao existirem, o script cria quando necessario.

## Colunas da Aba Geral

O sistema tenta localizar colunas por nomes equivalentes. Os campos principais sao:

- data de liberacao
- versao
- requisito
- cliente
- status
- caminho
- revisao

Datas sao aceitas como valor de data do Sheets, `dd/MM/yyyy` ou `yyyy-MM-dd`.

## Propriedades do Script

Configure em Apps Script > Project Settings > Script properties:

- `SPREADSHEET_ID`: obrigatorio apenas para script standalone. Em script vinculado a planilha, pode ficar vazio.
- `POST_SECRET`: segredo para importacao via `doPost`.
- `ALLOWED_USERS`: lista de e-mails separados por virgula. Em producao, configure esta propriedade para restringir o Web App. Quando vazio, o Web App nao restringe por e-mail.

Exemplo:

```text
ALLOWED_USERS=usuario1@empresa.com,usuario2@empresa.com
```

## Publicacao

1. Abra o projeto no Google Apps Script.
2. Adicione os arquivos `codigo.gs`, `Index.html`, `Styles.html`, `Scripts.html` e `appsscript.json`.
3. Configure as propriedades do script.
4. Publique como Web App.
5. Teste primeiro com acesso restrito antes de liberar para outros usuarios.

O manifesto permite acesso ao Web App por URL, mas as operacoes internas validam `ALLOWED_USERS` quando a propriedade esta configurada. O endpoint `doPost` usa `POST_SECRET` em vez da sessao do navegador.

## Importacao KPI via POST

Envie JSON com este formato:

```json
{
  "secret": "valor-do-POST_SECRET",
  "data": [
    {
      "ALT_VERSAO": "TOTAL GERAL",
      "ABERTA": 0,
      "ANDAMENTO": 0,
      "CORRIGINDO": 0,
      "CORRIGIDO": 0,
      "CONFERIDO": 0,
      "TOTAL": 0,
      "KPI_ABERTA_ANDAMENTO": 0,
      "KPI_CORRIGINDO_CORRIGIDO": 0,
      "KPI_CONFERIDO": 0,
      "SEMAFORO_ABERTA_ANDAMENTO": "VERDE",
      "SEMAFORO_CORRIGINDO_CORRIGIDO": "VERDE",
      "SEMAFORO_CONFERIDO": "VERDE"
    }
  ]
}
```

## Observacoes de Manutencao

- O schema atual de `KPI_Historico` usa 15 colunas e inclui `SESSION_ID`.
- Dados legados sao migrados automaticamente ao salvar novo KPI.
- O dashboard usa cache curto para reduzir leituras repetidas da planilha.
- O ponto de entrada correto do Apps Script e `Index.html`. Nao crie uma variante `index.html`, pois sistemas Windows tratam os dois nomes como o mesmo arquivo.
- Importacoes vazias sao rejeitadas. Valores numericos invalidos tambem sao rejeitados para evitar indicadores incorretos.

## Importacoes Manuais Adicionais

### Retorno de Requisitos

Cole um array JSON na area `KPI TESTER > Retorno de Requisitos`. Campos obrigatorios:

```json
[
  {
    "ALT_DTATST": "2026-06-01",
    "ALT_VERDDL": "48.02",
    "OPE_CONFER": "H75",
    "TOTAL_SEM_RETORNO": 0,
    "TOTAL_COM_RETORNO": 1,
    "TOTAL_COM_E_SEM_RETORNO": 1,
    "TOTAL_REQ_RETORNO": 1
  }
]
```

### Testes de Copias

Cole um array JSON na area `KPI TESTER > Testes de Copias`. Campos obrigatorios:

```json
[
  {
    "ALT_DTATST": "2026-06-01",
    "ALT_VERDDL": "48.02",
    "OPE_CONFER": "H75",
    "TOTAL": 1,
    "TOTAL_COM_TESTE": 1,
    "TOTAL_SEM_TESTE": 0
  }
]
```

### Cadastro de Operadores

Cole um array JSON na aba `Cadastro Operadores`. A importacao substitui o cadastro atual e rejeita listas vazias ou codigos duplicados:

```json
[
  {
    "OPE_LOGOPE": "H75",
    "OPE_DESCRI": "Nome do operador"
  }
]
```
