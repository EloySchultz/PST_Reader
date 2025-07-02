# Read_PST.py

A Python script to extract emails from a `.pst` (Outlook Personal Storage Table) file and export them to:

- A CSV file containing metadata like subject, sender, date, and truncated body
- Individual folders (one per email) with the full body saved as an HTML file

## Features

- Extracts all emails from a `.pst` file recursively
- Saves metadata to a CSV file (sorted by date, newest first)
- Saves full email bodies (HTML or plaintext) as `body.html` in unique subfolders
- Robust against encoding errors and corrupt entries

---

## 📦 Setup Instructions


### 0. Clone this repository
In the folder where you want to have the program, run
```
git clone https://github.com/EloySchultz/PST_Reader.git
```

### 1. Create a virtual environment

```bash
python3 -m venv venv
source venv/bin/activate       # On Windows: venv\Scripts\activate
```

### 2. Install required packages

```bash
pip install libpff-python
pip install tqdm
```

> ⚠️ `libpff-python` is a Python binding for the libpff library. On Linux/macOS, you may need to install system dependencies first via `apt`, `brew`, or `conda`.

> ⚠️ `tqdm` is used for displaying the progress bar during extraction.

---

## ▶️ Usage

```bash
python Read_PST.py -pst_path PATH_TO_PST_FILE -output_path OUTPUT_FOLDER
```

### Arguments

| Argument        | Description                                     |
|-----------------|-------------------------------------------------|
| `-pst_path`     | Path to the input `.pst` file                   |
| `-output_path`  | Folder where output will be written             |

---

## 🧾 Example

```bash
python Read_PST.py -pst_path "emails/inbox.pst" -output_path "output/"
```

This will:
- Create `output/inbox.csv`
- Create one folder per email inside `output/`, named by unique ID (UUID)
- In each folder, save a `body.html` file containing the full email content

---

## 🗂️ Output Structure

```
output/
├── inbox.csv
├── 3fa85f64-5717-4562-b3fc-2c963f66afa6/
│   └── body.html
├── b58a2e4e-d5a0-432e-8183-d92733c8b50d/
│   └── body.html
└── ...
```

### CSV Columns

- `email_id`: Unique UUID for each email (matches subfolder name)
- `date_received`: ISO timestamp (if available)
- `subject`: Email subject
- `sender`: Display name of sender
- `body`: Truncated plain-text body (500 chars max)

---

## 🛠 Notes

- Sometimes, data contains UTF16 characters which are NOT supported. So, the field will be set as "unreadable" instead.
- HTML bodies are preferred, but plaintext is used as fallback
- Attachments are currently not supported
---
