# DOCX Translation Studio

A desktop app that translates `.docx` files paragraph by paragraph via the OpenAI API, preserving formatting, inline objects, and character styles. Runs on **Windows, macOS, and Linux**.

> **Note:** The Windows taskbar progress indicator (PyTaskbar) is automatically disabled on non-Windows systems. All other features work cross-platform.

> The script was completely written by AI

---

## Requirements

- Python 3.11 or newer
- An OpenAI API key

---

## 1. Install Git (if not already installed)

Open a terminal and run:

**Windows (Bash):**
```bash
winget install --id Git.Git -e --source winget
```

**macOS:**
```applescript
brew install git
```

**Linux (Debian/Ubuntu):**
```bash
sudo apt install git
```

Restart the terminal after installation so `git` is available.

---

## 2. Install Python (if not already installed)

**Windows (Bash):**
```bash
winget install --id Python.Python.3.14 --source winget
```

**macOS:**
```bash
brew install python
```

**Linux (Debian/Ubuntu):**
```bash
sudo apt install python3 python3-pip python3-tk
```

Restart the terminal afterwards. Verify with:

```bat
python --version
```

---

## 3. Clone the repository

```bat
git clone https://github.com/wdev95/Docx-Translator-OpenAI
cd Docx-Translator-OpenAI
```

> **No Git?** Download the ZIP from the GitHub page (green *Code* button → *Download ZIP*), extract it, and open the folder in a terminal.

---

## 4. Install dependencies

All third-party packages can be installed with a single command:

```
pip install openai python-docx sv-ttk darkdetect PyTaskbar
```

> On macOS/Linux `PyTaskbar` is Windows-only — you can omit it:
> ```bash
> pip install openai python-docx sv-ttk darkdetect
> ```

| Package | Purpose |
|---|---|
| `openai` | OpenAI API client |
| `python-docx` | Read and write `.docx` files |
| `sv-ttk` | Modern Sun Valley UI theme for Tkinter |
| `darkdetect` | Detect Windows dark/light mode |
| `PyTaskbar` | Taskbar progress indicator for Windows 7+ |

---

## 5. Start the app

```bat
python translate.py
```

---

## Features

| Feature | Description |
|---|---|
| **API Key** | Enter your OpenAI API key in *Settings*. It is saved locally in `translate_settings.json`. |
| **Paragraph styles** | Click *Styles* to choose which Word paragraph styles should be translated (e.g. `Standard`, `Normal`). |
| **Character styles** | Click *Char Styles* to select inline character styles that carry formatting anchors (bold, color, etc.). |
| **Model selection** | Click *Load Models* to fetch available OpenAI models and select one. Pricing per 1 M tokens is shown inline. |
| **Glossary** | Add fixed term pairs (source → target) that the model will be instructed to respect. |
| **Translate** | Click *Start Translation*. Paragraphs are grouped by style and sent to the API in batches. Progress is shown in the progress bar and the Windows taskbar. |
| **Incremental update** | After a first successful translation a snapshot of the source file is saved. On subsequent runs only changed or added paragraphs are re-translated; unchanged blocks are left untouched in the output file. |
| **Pause / Cancel** | The translation can be paused or cancelled at any time. A checkpoint file is saved on pause. |
| **Formatting preservation** | Bold, color, inline formulas (OMML), and embedded objects are protected by anchor tokens and restored exactly in the output. |
