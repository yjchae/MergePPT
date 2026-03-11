# MergePPT 🎨

A dedicated Python-based GUI tool for merging multiple PowerPoint presentations effortlessly.

---

## ✨ Key Features

- **Format Support** — Supports both `.ppt` and `.pptx` files
- **Batch Merging** — Combine any number of files into a single presentation
- **User-Friendly GUI** — Built with PySide6 for a modern and clean interface
- **Preserve Layouts** — Maintains the original slide formatting and design
- **Drag & Drop** — Simply drag files into the window to add them
- **Reorder Freely** — Rearrange slides by dragging items in the list
- **Black Divider Slide** — Automatically inserts a black blank slide between each file
- **Cross-Platform** — Runs on macOS and Windows

---

## 📸 Preview

> _Dark-themed modern UI with drag-and-drop support_
<img width="681" height="611" alt="image" src="https://github.com/user-attachments/assets/744533ae-13ab-4f47-a89b-1dcca0b2f63f" />

---

## 🚀 Download

| Platform | Download |
|----------|----------|
| macOS    | [PPT병합기_mac.dmg](../../releases/latest) |
| Windows  | [PPT병합기.exe](../../releases/latest) |

---

## 🛠 Development Setup

### Requirements

- Python 3.11+
- [LibreOffice](https://www.libreoffice.org/) _(required for `.ppt` → `.pptx` conversion)_

### Install dependencies

```bash
pip install PySide6 python-pptx lxml
```

### Run

```bash
python mergeppt.py
```

---

## 📦 Build

### macOS

```bash
pip install pyinstaller
bash build_mac.sh
# Output: dist/PPT병합기_mac.dmg
```

### Windows

```batch
pip install pyinstaller
build_win.bat
:: Output: dist/PPT병합기.exe
```

### Automated build via GitHub Actions

Push a version tag to trigger automatic builds for both platforms:

```bash
git tag v1.0
git push origin v1.0
```

Artifacts will be available on the [Releases](../../releases) page.

---

## 📋 How to Use

1. Launch the app
2. Drag and drop `.ppt` / `.pptx` files into the window
3. Reorder files by dragging items in the list
4. Click **최종 파일로 합치기** to merge
5. Choose a save location — done!

> `.ppt` files are automatically converted to `.pptx` via LibreOffice before merging.

---

## 🗂 Project Structure

```
MergePPT/
├── mergeppt.py        # Main application
├── PPTMerger.spec     # PyInstaller spec file
├── build_mac.sh       # macOS build script
├── build_win.bat      # Windows build script
└── .github/
    └── workflows/
        └── build.yml  # GitHub Actions CI/CD
```

---

## 📄 License

MIT License © 2026 @ZionP
