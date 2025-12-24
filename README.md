# 📊 PPT to Markdown Converter

A tool that automatically converts PowerPoint files (.ppt, .pptx) into Markdown format. All slides are combined into a single Markdown file. Available in both GUI and command-line versions.

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![PyQt6](https://img.shields.io/badge/PyQt6-6.0+-green.svg)](https://www.riverbankcomputing.com/static/Docs/PyQt6/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## ✨ Features

- 📄 **PowerPoint File Reading**: Automatically recognizes all slides in .ppt and .pptx format files
- 📝 **Markdown Conversion**: Converts all slides into a single Markdown file with proper formatting
- 🖱️ **Drag and Drop**: Easily convert files by dragging them in the GUI
- 📊 **Real-time Progress**: Monitor conversion progress in real-time
- 🔄 **Asynchronous Processing**: GUI remains responsive during conversion
- 📦 **Executable Support**: Generate .exe files that run without Python installation
- 📋 **Content Extraction**: Extracts text, tables, lists, and image placeholders from slides

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone <repository-url>
cd ppt2md.git

# Create and activate virtual environment
python -m venv venv
venv\Scripts\activate  # Windows

# Install dependencies
pip install python-pptx PyQt6
```

### Usage

#### GUI Version (Recommended)

```bash
python ppt2md_gui.py
```

1. When the GUI window opens, drag and drop a PPT/PPTX file or click the "Select File" button
2. Monitor the conversion progress
3. After conversion completes, a Markdown file is created in the same location as the source file

#### Command-Line Version

```bash
# Basic usage
python ppt2md.py "path/to/presentation.pptx"

# Specify output directory
python ppt2md.py "path/to/presentation.pptx" "output/path"
```

## 📋 Requirements

- Python 3.8 or higher
- Windows 10 or higher (GUI version)

### Required Packages

```
python-pptx>=0.6.0
PyQt6>=6.0.0
```

## 📁 Project Structure

```
ppt2md.git/
├── ppt2md.py              # Command-line version (core conversion logic)
├── ppt2md_gui.py          # GUI version
├── build_exe.bat          # Windows batch build script
├── build_exe.ps1          # PowerShell build script
├── README_BUILD.md        # Build guide
└── doc/                   # PowerPoint files to convert
```

## 🔧 Build and Deployment

### Creating Executable (.exe) File

```bash
# Using batch file
build_exe.bat

# Or using PowerShell script
.\build_exe.ps1
```

After build completes, the `dist\PPT2Markdown.exe` file will be created.

For detailed build instructions, please refer to [README_BUILD.md](README_BUILD.md).

## 📖 Usage Examples

### Input File
- `presentation.pptx` (contains multiple slides)

### Output File
A single Markdown file containing all slides:
- `presentation.md`

### Markdown File Format

```markdown
# Presentation Name
*Source file: presentation.pptx*
*Total slides: 5*
---

## Slide 1

### Title

Content text here
- Bullet point 1
- Bullet point 2

---

## Slide 2

| Column1 | Column2 | Column3 |
|---------|---------|---------|
| Data1   | Data2   | Data3   |

[Image: Slide 2]

---
```

## 🛠️ Tech Stack

- **Python 3.8+**: Programming language
- **PyQt6**: GUI framework
- **python-pptx**: PowerPoint file reading and parsing
- **PyInstaller**: Executable build tool

## 📝 Key Features Details

### Content Extraction
- **Text**: Extracts text from text boxes and shapes
- **Tables**: Converts PowerPoint tables to Markdown tables
- **Lists**: Preserves bullet points and numbered lists
- **Images**: Replaces images with placeholder text `[Image: Slide N]`
- **Titles**: Automatically detects and formats slide titles

### Special Character Handling
- Pipe character (`|`): Escaped as `\|` in tables
- Line breaks: Preserved in text content
- Empty slides: Marked with "*This slide is empty.*"

### Error Handling
- File existence verification
- Empty slide processing
- Continues processing even if individual slide processing fails
- Provides detailed error messages and tracebacks

## 🤝 Contributing

Bug reports, feature suggestions, and Pull Requests are welcome!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is distributed under the MIT License.

## 🙏 Acknowledgments

This project was developed to convert PowerPoint-based presentation documents into Markdown format, making it easier to integrate with version control systems and documentation workflows.

## 📚 Related Documentation

- [Build Guide](README_BUILD.md)
- [한국어 문서](README_KR.md)

---

**Made with ❤️ for better documentation workflow**

