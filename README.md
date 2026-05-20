```markdown
# 📄 PDF Translator — Streamlit App

AI-powered semi-automatic book translation pipeline with EPUB/DOCX export.

![Main UI](Main%20UI.png)

---

## 🚀 Quick Start

### 1. Install dependencies

```
pip install -r requirements.txt
```

### 2. Run the app

```
streamlit run app.py
```

The app will open in your browser at **http://localhost:8501**

---

## 🛠 How to Use

1. **Upload a PDF** — use the button at the top of the page
2. **Select a fragment** in the sidebar on the left
3. **Click "Copy prompt"** and paste it into ChatGPT or Claude
4. **Paste the translation** into the "Paste translation here" field
5. **Click "Save translation"** — the app will automatically move to the next fragment
6. Once all fragments are translated — **download as JSON or DOCX**

---

## 📦 Dependencies

| Library       | Purpose                  |
| ------------- | ------------------------ |
| `streamlit`   | Web interface            |
| `pymupdf`     | PDF text extraction      |
| `python-docx` | Export to Word (.docx)   |

---

## 📂 Project Structure

```
pdf_translator/
├── app.py            # Main application code
├── requirements.txt  # Dependencies
└── README.md         # This file
```

---
---

# 📄 PDF Translator — Streamlit App

Приложение для поэтапного перевода PDF-документов с помощью ChatGPT / Claude.

---

## 🚀 Быстрый старт

### 1. Установи зависимости

```
pip install -r requirements.txt
```

### 2. Запусти приложение

```
streamlit run app.py
```

Откроется браузер по адресу **http://localhost:8501**

---

## 🛠 Как пользоваться

1. **Загрузи PDF** — кнопка вверху страницы
2. **Выбери фрагмент** в боковой панели (sidebar слева)
3. **Нажми «Копировать промпт»** и вставь его в ChatGPT или Claude
4. **Вставь полученный перевод** в поле «Вставь перевод сюда»
5. **Нажми «Сохранить перевод»** — приложение автоматически перейдёт к следующему фрагменту
6. Когда все фрагменты переведены — **скачай JSON или DOCX**

---

## 📦 Зависимости

| Библиотека    | Для чего                 |
| ------------- | ------------------------ |
| `streamlit`   | Веб-интерфейс            |
| `pymupdf`     | Извлечение текста из PDF |
| `python-docx` | Экспорт в Word (.docx)   |

---

## 📂 Структура проекта

```
pdf_translator/
├── app.py            # Основной код приложения
├── requirements.txt  # Зависимости
└── README.md         # Эта инструкция
```
```
