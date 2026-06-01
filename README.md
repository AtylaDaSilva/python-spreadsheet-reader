# Leitor Planilha Python

Biblioteca utilitária em Python para ler, modificar e salvar planilhas **Excel (.xlsx)**. O núcleo do projeto é a classe `LeitorXLSX`, que encapsula o [openpyxl](https://openpyxl.readthedocs.io/) e oferece leitura linha a linha, acesso a células, inserção de imagens, detecção de arquivo bloqueado (Excel ou LibreOffice) e formatação básica via `EstilizadorDeCelula`.

## Requisitos

- Python 3.10+ (o código usa `match`/`case` e união de tipos com `|`)
- Dependências listadas em `requirements.txt`:
  - `openpyxl` — leitura/escrita de `.xlsx`
  - `Pillow` — suporte a imagens no openpyxl

## Como rodar localmente

### 1. Clonar o repositório e entrar na pasta

```bash
cd leitor-planilha-python
```

### 2. Criar ambiente virtual (recomendado)

```bash
python -m venv .venv
```

**Windows (PowerShell):**

```powershell
.\.venv\Scripts\Activate.ps1
```

**Linux/macOS:**

```bash
source .venv/bin/activate
```

### 3. Instalar dependências

```bash
pip install -r requirements.txt
```

### 4. Usar no seu código

O módulo principal está em `main.py`. Importe as classes e use conforme o exemplo:

```python
from main import LeitorXLSX, EstilizadorDeCelula

leitor = LeitorXLSX("caminho/para/planilha.xlsx")

# Ler todas as linhas da planilha ativa
dados = leitor.ler_excel()
# dados é dict[int, tuple]: chave = número da linha (1-based), valor = tupla de valores da linha

valor = leitor.get_valor_celula(linha=1, coluna=1)
leitor.set_valor_celula(linha=2, coluna=1, valor="Novo texto")

celula = leitor.get_celula(linha=1, coluna=1)
EstilizadorDeCelula(celula).alinhar("center", "center", quebrar_texo=True)

leitor.adicionar_imagem("logo.png", celula="A1", altura=50, largura=100)
leitor.salvar_planilha()  # salva no caminho original
# leitor.salvar_planilha("saida/copia.xlsx")  # salva em outro caminho

leitor.fechar_planilha()
```

> **Nota:** Feche a planilha no Excel/LibreOffice antes de abrir com este leitor. Se o arquivo estiver em uso, `SpreadsheetIsLockedException` será lançada na abertura.

---

## Detalhamento dos objetos

### Classes

#### `LeitorXLSX`

Ponto de entrada para trabalhar com um arquivo `.xlsx`. Mantém o workbook em memória (`openpyxl.Workbook`) e cache dos dados lidos em `_sheet_data`.

| Membro | Descrição |
|--------|-----------|
| `caminho_planilha` | `Path` do arquivo informado no construtor. |
| `_sheet_data` | `dict[int, tuple]` com o resultado da última leitura via `ler_excel()`. |
| `_workbook` | Instância do workbook openpyxl ou `None` se fechado. |

**Construtor**

```python
LeitorXLSX(caminho_planilha: str | Path, *args, **kwargs)
```

- `caminho_planilha`: caminho para o arquivo `.xlsx`.
- `*args`, `**kwargs`: repassados para `openpyxl.load_workbook()` (ex.: `read_only=True`, `data_only=True`).
- Na inicialização, tenta abrir o workbook; se existir arquivo de lock do Excel/LibreOffice, levanta `SpreadsheetIsLockedException`.

**API pública**

| Método | Retorno | Descrição |
|--------|---------|-----------|
| `ler_excel()` | `dict[int, tuple]` | Valida existência e extensão `.xlsx`, lê todas as linhas da **planilha ativa** e retorna um dicionário cuja chave é o índice da linha (começando em 1) e o valor é uma tupla com os valores das colunas. |
| `get_celula(linha, coluna)` | `Cell \| MergedCell` | Retorna o objeto de célula do openpyxl (linhas e colunas são **1-based**). |
| `get_valor_celula(linha, coluna)` | `Any` | Atalho para `.value` da célula. |
| `set_valor_celula(linha, coluna, valor)` | `None` | Define o valor da célula. |
| `adicionar_imagem(caminho_imagem, celula, altura, largura)` | — | Insere imagem na planilha ativa; `celula` é referência estilo Excel (ex.: `"B2"`). |
| `salvar_planilha(caminho=None)` | `Path` | Persiste o workbook. Se `caminho` for omitido, usa `caminho_planilha`. Cria diretórios pais se necessário. |
| `fechar_planilha()` | — | Fecha o workbook e limpa `_sheet_data`. |
| `is_spreadsheet_locked()` | `bool` | Verifica se há arquivo de lock do Excel (`~$nome.xlsx`) ou LibreOffice (`.~lock.nome.xlsx#`) no mesmo diretório. |

**Métodos protegidos (uso interno / extensão)**

| Método | Descrição |
|--------|-----------|
| `_abrir_workbook(*args, **kwargs)` | Carrega ou reutiliza o workbook; falha se a planilha estiver bloqueada. |
| `_fechar_workbook()` | Fecha e zera estado interno. |
| `_planilha_ativa()` | Retorna a worksheet ativa; levanta `NoActiveSheetException` se não houver. |
| `_read_xlsx()` | Itera `iter_rows(values_only=True)` e popula `_sheet_data`. |
| `_is_xlsx_locked()` | Detecta lock temporário do Excel. |
| `_is_ods_locked()` | Detecta lock do LibreOffice (padrão `.~lock.<arquivo>#`). |

**Exceções que podem ser lançadas (além das customizadas)**

- `FileNotFoundError`: arquivo inexistente em `ler_excel()` ou `is_spreadsheet_locked()`.
- `ValueError`: extensão diferente de `.xlsx` em `ler_excel()`.

---

#### `EstilizadorDeCelula`

Helper para aplicar estilo a uma célula já obtida via `LeitorXLSX.get_celula()`.

| Membro | Descrição |
|--------|-----------|
| `celula` | Referência `Cell` ou `MergedCell` do openpyxl. |

**Construtor**

```python
EstilizadorDeCelula(celula: Cell | MergedCell)
```

**Métodos**

| Método | Descrição |
|--------|-----------|
| `alinhar(horizontal, vertical, quebrar_texo)` | Define `Alignment` (valores aceitos pelo openpyxl, ex.: `"left"`, `"center"`, `"right"`). `quebrar_texo` mapeia para `wrapText`. |
| `formatar(formato)` | Define `number_format` da célula (ex.: `"#,##0.00"`, `"dd/mm/yyyy"`). |

---

### Exceções

#### `NoActiveSheetException`

Herda de `Exception`. Levantada em `_planilha_ativa()` quando o workbook não possui planilha ativa (`active` é `None`). Indica arquivo corrompido, vazio ou estrutura inesperada.

#### `SpreadsheetIsLockedException`

Herda de `Exception`. Levantada em `_abrir_workbook()` quando `is_spreadsheet_locked()` retorna `True`, ou seja, outro processo (Excel ou LibreOffice) mantém o arquivo aberto. Evita leitura/gravação concorrente que poderia corromper dados ou falhar silenciosamente.

---

## Estrutura do repositório

```
leitor-planilha-python/
├── main.py           # LeitorXLSX, EstilizadorDeCelula e exceções
├── requirements.txt
└── README.md
```

## Licença

Consulte o repositório ou os mantenedores do projeto para informações de licenciamento.
