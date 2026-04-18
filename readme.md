# pdfcraft by mrchn (студия мрачно)

[english](#english) | [русский](#русский)

<a name='english'></a>

**pdfcraft** is a lightweight and powerful python3 library designed for automated generation of pdf documents from microsoft word (.docx) templates. it handles templating, conversion, and batch processing, making it ideal for automating contracts, invoices, and reports.

### setup
```
pip install pdfcraft-mrchn
```

### features
* **template rendering:** uses `docxtpl` to populate word (.docx) documents with dynamic data.
* **flexible conversion:** supports both `docx2pdf` (windows-optimized) and libreoffice (cross-platform).
* **batch processing:** handles large datasets with parallel execution (multiprocessing).
* **robustness:** includes data validation, error handling, and dry-run mode.
* **archival-ready:** supports PDF/A-1b export for long-term storage.

### quick start
```python
from pdfcraft import Generate

generator = Generate(logging=True) # initialize

# generate a single document
generator.process({'name': 'John Doe', 'id': '001'}, out_path='contract_001.pdf')

# batch process
data_list = [{'id': '001', 'name': 'John'}, {'id': '002', 'name': 'Jane'}]
generator.batch_process(data_list, out_dir='output')
```

<a name='русский'></a>

**pdfcraft** — это лёгкая и мощная библиотека python3 для автоматической генерации pdf-документов из шаблонов microsoft word (.docx). она берёт на себя заполнение шаблонов, конвертацию и массовую обработку, что делает её идеальной для создания договоров, счетов и отчетов.

### установка
```
pip install pdfcraft-mrchn
```

### возможности
* **заполнение шаблонов:** использует `docxtpl` для вставки динамических данных в документы word (.docx).
* **гибкая конвертация:** поддержка `docx2pdf` (для windows) и libreoffice (кроссплатформенно).
* **массовая обработка:** обработка больших наборов данных с поддержкой параллельных вычислений (multiprocessing).
* **надежность:** валидация данных, обработка ошибок и режим тестирования (dry-run).
* **архивное качество:** поддержка экспорта в формат PDF/A-1b.

### начало работы
```python
from pdfcraft import Generate

generator = Generate(logging=True) # инициализация

# генерация одного документа
generator.process({'name': 'Иван Иванов', 'id': '001'}, out_path='contract_001.pdf')

# массовая генерация
data_list = [{'id': '001', 'name': 'Иван'}, {'id': '002', 'name': 'Олег'}]
generator.batch_process(data_list, out_dir='output')
```