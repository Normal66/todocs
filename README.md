# todocs
Генерация Word (.docx) из Excel: по образцу с жёлтой подсветкой строятся плейсхолдеры по заголовкам столбцов; дата, основание и МОЛ — в DOP.JSON; универсальный CLI (--word, --excel); в готовых файлах подсветка снимается. Локальный инструмент, Python.

## todocs — Word (DOCX) из Excel по шаблону (RU/EN)

### Русский

`todocs` генерирует Word-файлы (`.docx`) из Excel (`.xlsx`) по образцу Word:

- ты даёшь **исходный Word** (например `Образец акт.docx`), где нужные значения выделены **жёлтой подсветкой**;
- утилита находит все жёлтые фрагменты и превращает их в плейсхолдеры `{{...}}`;
- имена плейсхолдеров берутся из **названий столбцов** первой строки Excel (приводятся к безопасному виду);
- затем для каждой позиции Excel создаётся отдельный `.docx`;
- на выходе **подсветка убирается** (цветом ничего не выделяется).

#### Требования

- Windows 10
- Python 3.11+ (или существующий `venv` в проекте)

#### Установка зависимостей

```powershell
Set-Location f:\Develop\Palata\todocs
.\venv\Scripts\python.exe -m pip install -U pip
.\venv\Scripts\python.exe -m pip install openpyxl docxtpl
```

#### Конфиг DOP.JSON (опционально)

Файл `DOP.JSON` задаёт значения, которые не берутся из Excel:

- `ACT_DATE`
- `BASIS`
- `MOL`

Если `DOP.JSON` отсутствует или поле пустое — используется дефолт (как в примере `660001361250.docx`).

#### Запуск (универсальный)

```powershell
Set-Location f:\Develop\Palata\todocs
.\venv\Scripts\python.exe .\tools\todocs_cli.py --word ".\Образец акт.docx" --excel ".\Инв-сер.xlsx" --out ".\out"
```

Дополнительно:

- сохранить шаблон в конкретный файл:

```powershell
.\venv\Scripts\python.exe .\tools\todocs_cli.py --word ".\Образец акт.docx" --excel ".\Инв-сер.xlsx" --template-out ".\template_act.docx"
```

- указать другой `DOP.JSON`:

```powershell
.\venv\Scripts\python.exe .\tools\todocs_cli.py --word ".\Образец акт.docx" --excel ".\Инв-сер.xlsx" --dop ".\DOP.JSON"
```

Результат: `out\<инвентарный номер>.docx` (например `out\660001361250.docx`).

---

### English

`todocs` generates Word documents (`.docx`) from an Excel file (`.xlsx`) using a Word sample:

- you provide a **source Word sample** (e.g. `Образец акт.docx`) where the fields to be replaced are **yellow-highlighted**;
- the tool finds all yellow-highlighted fragments and converts them into `{{...}}` placeholders;
- placeholder names are derived from the **Excel header row** (sanitized into safe identifiers);
- it produces one output `.docx` per Excel item;
- the output documents have **no highlight coloring** (highlight is cleared).

#### Requirements

- Windows 10
- Python 3.11+ (or the existing `venv` in this repo)

#### Install dependencies

```powershell
Set-Location f:\Develop\Palata\todocs
.\venv\Scripts\python.exe -m pip install -U pip
.\venv\Scripts\python.exe -m pip install openpyxl docxtpl
```

#### Optional DOP.JSON

`DOP.JSON` provides values not coming from Excel:

- `ACT_DATE`
- `BASIS`
- `MOL`

If `DOP.JSON` is missing or a field is empty, defaults are used (matching the sample `660001361250.docx`).

#### Run (universal)

```powershell
Set-Location f:\Develop\Palata\todocs
.\venv\Scripts\python.exe .\tools\todocs_cli.py --word ".\Образец акт.docx" --excel ".\Инв-сер.xlsx" --out ".\out"
```


