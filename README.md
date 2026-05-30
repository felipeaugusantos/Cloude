# Sistema de Gestao de Copias

Dashboard em Google Apps Script para acompanhar liberacoes de copias, relatorios operacionais e KPI de fechamento.

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
- `ALLOWED_USERS`: opcional. Lista de e-mails separados por virgula. Quando vazio, o Web App nao restringe por e-mail.

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
- Evite recriar o arquivo `index.html`; o ponto de entrada correto do Apps Script e `Index.html`.
