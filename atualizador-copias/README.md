# Atualizador de Cópias - Autpas

Ferramenta interna (protótipo) para o time de teste. Centraliza numa **única tela**:

1. **Atualizar as cópias** de vários testers de uma vez, rodando o `.bat` de cada pasta `Autpas` — sem precisar entrar pasta por pasta.
2. **Abrir o Autcom** (`autcom.exe`) direto pelo programa, dentro da pasta do tester escolhido.

## Como funciona

- O Jenkins compila a última versão numa pasta de saída (de hora em hora, ~35 min).
- Cada tester tem uma pasta própria `C:\<usuario>\Autpas` (conteúdo igual para todos).
- O `.bat` de cada pasta atualiza a `Autpas` com a última compilação.
- O programa apenas **orquestra**: dispara os `.bat` selecionados e abre o `autcom.exe`.

Nada da lógica atual muda — os `.bat` que já existem continuam sendo usados. A ferramenta só evita o trabalho manual de abrir pasta por pasta.

## Configuração

Copie `config.example.json` para `config.json` (na mesma pasta do `.exe`/script) e ajuste:

```json
{
  "autcom_exe": "autcom.exe",
  "bat_nome": "atualizar.bat",
  "usuarios": [
    { "nome": "Felipe.santos", "pasta": "C:\\Felipe.santos\\Autpas" },
    { "nome": "Aline.pizzo",   "pasta": "C:\\Aline.pizzo\\Autpas" }
  ]
}
```

- `bat_nome`: nome do `.bat` dentro de cada pasta `Autpas`. Por padrão o programa roda `<pasta>\atualizar.bat`.
- `autcom_exe`: nome do executável do Autcom dentro da pasta.
- Cada usuário pode sobrescrever o caminho do `.bat` ou do Autcom com os campos opcionais `"bat"` e `"autcom"`.

## Rodar em desenvolvimento

Requer Python 3.9+ (o Tkinter já vem no Python padrão do Windows).

```bash
python atualizador.py
```

## Gerar o `.exe` único

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --name AtualizadorCopias atualizador.py
```

O executável sai em `dist/AtualizadorCopias.exe`. Distribua junto com o `config.json`.

## Status

Protótipo inicial para validação com o time. Pontos que ainda podem evoluir:

- Detectar automaticamente as pastas de usuário (varrer `C:\` em vez de listar no config).
- Mostrar a hora da última compilação do Jenkins / avisar se a build está em andamento.
- Botão "Abrir Autcom" por linha, com atualização + abertura em um clique só.
