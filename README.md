# Byte Hoover 🌀

**Professional Data Cleaning Tool for the Web**

Byte Hoover is a powerful, browser-based data cleaning application that transforms messy spreadsheets into clean, structured data. Built with vanilla JavaScript and Bootstrap, it runs entirely in your browser—no server, no uploads to external services.

## ✨ Features

- **Smart Data Cleaning**: Remove duplicates, trim whitespace, fix punctuation, validate emails/phones
- **Intelligent Formatting**: Smart capitalization, date standardization, name formatting
- **Multiple File Support**: Process Excel (.xlsx, .xls), CSV, and text files
- **Batch Processing**: Clean multiple files at once
- **Export Options**: Save as XLSX, CSV, or JSON with optional ZIP compression
- **Preset Configurations**: Quick-start templates for common use cases
- **Real-time Preview**: See before/after comparisons
- **Undo/Redo**: Full history tracking of cleaning actions
- **Dark/Light Mode**: Choose your preferred theme
- **100% Client-side**: Your data never leaves your browser

## 🚀 Quick Start

1. **Upload your file** (drag & drop or click to browse)
2. **Choose cleaning options** or select a preset
3. **Click "Inhale & Cleanse Data"**
4. **Download your cleaned file**

No installation required! Just open `index.html` in any modern browser.

## 📁 Project Structure
- **byte-hoover/
- **|
- **|__css/
- **|  |
- **|  |__style.css
- **|__js/
- **|  |
- **|__|__script.js
- **|
- **|__index.html

## 🔧 Supported Operations

- **Capitalization**: Smart title case with name recognition
- **Punctuation**: Clean excessive punctuation
- **Spacing**: Remove extra spaces, preserve single spaces
- **Phone Numbers**: Format and validate phone numbers
- **Dates**: Standardize date formats
- **Emails**: Validate and lowercase emails
- **Empty Rows/Columns**: Remove completely empty rows and columns
- **Duplicates**: Remove duplicate rows

## 🎯 Use Cases

- **Marketing**: Clean contact lists before email campaigns
- **Research**: Standardize survey data for analysis
- **HR**: Prepare employee databases for migration
- **Finance**: Sanitize transaction records for reporting
- **Education**: Format student data for systems integration

## 🎨 Presets

1. **Basic Clean**: Essential formatting fixes
2. **Contact List**: Optimized for email/phone databases
3. **Export Ready**: Comprehensive cleaning for system imports
4. **Sensitive Data**: Includes anonymization options

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl/Cmd + O` | Open file dialog |
| `Ctrl/Cmd + Enter` | Process files |
| `Ctrl/Cmd + Z` | Undo |
| `Ctrl/Cmd + Shift + Z` | Redo |
| `Ctrl/Cmd + P` | Preview data |
| `Ctrl/Cmd + D` | Toggle dark mode |

## 🔒 Privacy & Security

- **Zero Data Transmission**: Everything processes locally
- **No Tracking**: No analytics, no cookies, no telemetry
- **Transparent Code**: Open source and inspectable
- **Temporary Storage**: Files only exist in memory during processing

## 🛡️ License

MIT License

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Submit a pull request

---

**Byte Hoover**: Because clean data shouldn't require a data science degree. 🧹✨
