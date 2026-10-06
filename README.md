# AI Document & Presentation Skills Suite

Профессиональный набор локальных AI-скиллов для анализа документов (`.docx`, `.pdf`) и создания презентаций PowerPoint (`.pptx`) **без внешних платных API и без передачи данных в облако**.

Все скиллы соответствуют открытому стандарту **Agent Skills (`SKILL.md`)** и автоматически распознаются AI-ассистентами (Google Antigravity, Claude Code, Cursor, Copilot) благодаря расположению в `.agents/skills/` и `skills/`.

---

## 📦 Каталог установленных AI-скиллов (8 скиллов)

### 📄 Чтение, парсинг и анализ документов (DOCX и PDF)

| Скилл | Назначение | Ключевые возможности |
|---|---|---|
| **[`docx-reader-analyzer`](./skills/docx-reader-analyzer/SKILL.md)** | Анализ и чтение Word (`.docx`, `.dotx`) | Извлечение структуры заголовков (Title, H1–H3), параграфов, метаданных и таблиц в формате Markdown. Экспорт всего документа в чистый Markdown для анализа AI. Работает через `python-docx`. |
| **[`pdf-reader-analyzer`](./skills/pdf-reader-analyzer/SKILL.md)** | Анализ и чтение PDF (`.pdf`) | Постраничное извлечение текста с сохранением колонок и отступов, точное извлечение финансовых/технических таблиц в Markdown-формат, поиск по страницам, извлечение метаданных. Работает через `pypdf` и `pdfplumber`. |

### 📊 Создание и архитектура презентаций (PPTX)

| Скилл | Назначение | Ключевые возможности |
|---|---|---|
| **[`presentation-architect`](./skills/presentation-architect/SKILL.md)** | Архитектура и сторителлинг | Формула Pixar (Story Spine), принцип пирамиды Минто (SCQA), каталог 12 типов слайдов, Action Titles (утвердительные заголовки), микрокопия и структура заметок докладчика. |
| **[`python-pptx-pro`](./skills/python-pptx-pro/SKILL.md)** | Локальный Python-генератор PPTX | Создание широкоформатных слайдов 16:9 на чистом Python (`python-pptx`). Готовый модуль `pptx_engine.py`: карточки, KPI-метрики, матрицы сравнения, таймлайны, нативные редактируемые графики. 3 дизайн-палитры (`corporate`, `dark_tech`, `vibrant`). |
| **[`presentation-pitch-deck`](./skills/presentation-pitch-deck/SKILL.md)** | Питч-деки для инвесторов и бизнеса | Канонический 10-слайдовый фреймворк Sequoia Capital / Y Combinator / Guy Kawasaki. Расчет TAM/SAM/SOM, юнит-экономика SaaS, слайд "Why Now". |
| **[`pptx-ooxml-master`](./skills/pptx-ooxml-master/SKILL.md)** | Низкоуровневый OOXML | Распаковка `.pptx`, дублирование слайдов и макетов существующих шаблонов через `manage_slide.py` с сохранением связей Open Packaging Conventions. |
| **[`marp-pptx`](./skills/marp-pptx/SKILL.md)** | Конвертация Markdown в PPTX | Мгновенный экспорт структурированного Markdown в нативный PowerPoint через `npx @marp-team/marp-cli`. Стилизация CSS, многоколоночные сетки. |
| **[`pptx-qa-inspector`](./skills/pptx-qa-inspector/SKILL.md)** | Автоматический QA-аудит | Скрипт `audit_presentation.py` проверяет плотность текста (слов на слайд), соотношение сторон 16:9, покрытие заметками спикера и пустые элементы. |

---

## 🔄 Сквозной рабочий процесс: из DOCX / PDF в готовую презентацию PPTX

Когда у вас есть отчет, аналитическая записка, регламент или whitepaper, цепочка обработки строится автоматически:

```text
[DOCX / PDF документ]
       │
       ▼  (docx-reader-analyzer / pdf-reader-analyzer)
[Структурированный Markdown + Таблицы + Метаданные]
       │
       ▼  (presentation-architect)
[Story Spine + Слайд-план + Action Titles + Заметки докладчика]
       │
       ▼  (python-pptx-pro)
[Готовая презентация 16:9 .pptx со стилями и карточками]
       │
       ▼  (pptx-qa-inspector)
[Отчет контроля качества: 100% готовность к показу]
```

---

## 🚀 Быстрый старт: Команды в терминале

### 1. Анализ DOCX документа:
```powershell
# Общая структура и оглавление
python .agents/skills/docx-reader-analyzer/scripts/read_docx.py document.docx

# Экспорт в Markdown (для подачи в AI)
python .agents/skills/docx-reader-analyzer/scripts/read_docx.py document.docx --markdown > document.md

# Извлечение только таблиц
python .agents/skills/docx-reader-analyzer/scripts/read_docx.py document.docx --tables
```

### 2. Анализ PDF документа:
```powershell
# Общая сводка и число страниц/слов
python .agents/skills/pdf-reader-analyzer/scripts/read_pdf.py document.pdf

# Экспорт страниц и таблиц в Markdown
python .agents/skills/pdf-reader-analyzer/scripts/read_pdf.py document.pdf --markdown > document.md

# Извлечение таблиц
python .agents/skills/pdf-reader-analyzer/scripts/read_pdf.py document.pdf --tables
```

### 3. Генерация презентации PowerPoint:
```powershell
python .agents/skills/python-pptx-pro/scripts/generate_demo_deck.py my_presentation.pptx
```

### 4. Проверка качества сгенерированного файла:
```powershell
python .agents/skills/pptx-qa-inspector/scripts/audit_presentation.py my_presentation.pptx
```

---

## 🔒 100% Offline и приватность
- Все скрипты и библиотеки (`python-docx`, `pypdf`, `pdfplumber`, `python-pptx`) установлены в окружении и выполняются локально.
- Никаких сетевых запросов, утечек корпоративных документов или платных SaaS-ключей.
