---
name: "md2docx"
description: "Converts Markdown files to Word documents. Invoke when user needs to convert .md files to .docx format."
---

# Markdown to Word Converter

This skill converts Markdown (.md) files to Word (.docx) documents.

## Usage

### Command Format
```bash
python md2docx.py <input.md> <output.docx>
```

### Parameters
- `<input.md>`: Path to the input Markdown file
- `<output.docx>`: Path to the output Word document

### Examples

#### Example 1: Basic conversion
```bash
python md2docx.py document.md output.docx
```

#### Example 2: Convert with custom output path
```bash
python md2docx.py docs/README.md docs/output/document.docx
```

## Features
- Supports basic Markdown formatting
- Automatically creates output directories if they don't exist
- Provides clear error messages for missing files

## Dependencies
- Python 3.6+
- markdown library
- python-docx library
- lxml library

## Installation
```bash
python -m pip install markdown python-docx lxml
```
