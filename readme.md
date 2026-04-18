# pdfcraft

[english](#english) | [русский](#русский)

<a name='english'></a>

**pdfcraft** is a lightweight and powerful python3 library designed for automated generation of pdf documents from Microsoft Word (.docx) templates. It handles templating, conversion, and batch processing, making it ideal for automating contracts, invoices, and reports.

### SETUP
```
pip install pdfcraft-mrchn
```

### FEATURES
* **Template Rendering:** Uses `docxtpl` to populate Word documents with dynamic data.
* **Flexible Conversion:** Supports both `docx2pdf` (Windows-optimized) and LibreOffice (cross-platform).
* **Batch Processing:** Handles large datasets with parallel execution (multiprocessing).
* **Robustness:** Includes data validation, error handling, and dry-run mode.
* **Archival Ready:** Supports PDF/A-1b export for long-term storage.

### QUICK START
```python
from pdfcraft import Generate

# initialize
generator = Generate(logging=True)

# generate a single document
generator.process({'name': 'John Doe', 'id': '001'}, out_path='contract_001.pdf')

# batch process
data_list = [{'id': '001', 'name': 'John'}, {'id': '002', 'name': 'Jane'}]
generator.batch_process(data_list, out_dir='output')
```

<a name='русский'></a>

**pdfcraft** — это лёгкая и мощная библиотека python3 для автоматической генерации pdf-документов из шаблонов Microsoft Word (.docx). Она берёт на себя заполнение шаблонов, конвертацию и массовую обработку, что делает её идеальной для создания договоров, счетов и отчетов.

### УСТАНОВКА
```
pip install pdfcraft-mrchn
```

### ВОЗМОЖНОСТИ
* **Заполнение шаблонов:** Использует `docxtpl` для вставки динамических данных в документы Word.
* **Гибкая конвертация:** Поддержка `docx2pdf` (для Windows) и LibreOffice (кроссплатформенно).
* **Массовая обработка:** Обработка больших наборов данных с поддержкой параллельных вычислений (multiprocessing).
* **Надежность:** Валидация данных, обработка ошибок и режим тестирования (dry-run).
* **Архивное качество:** Поддержка экспорта в формат PDF/A-1b.

### БЫСТРЫЙ СТАРТ
```python
from pdfcraft import Generate

generator = Generate(logging=True) # инициализация

# генерация одного документа
generator.process({'name': 'Иван Иванов', 'id': '001'}, out_path='contract_001.pdf')

# массовая генерация
data_list = [{'id': '001', 'name': 'Иван'}, {'id': '002', 'name': 'Олег'}]
generator.batch_process(data_list, out_dir='output')
```