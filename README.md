# Excel GPT Translator

A Python application that translates Excel files using GPT API while preserving the original formatting. This tool is designed for efficient and accurate translation of Excel documents, with support for field-specific terminology and customizable translation prompts.

## Features

- 🌐 Support for multiple languages
- 📊 Preserves Excel formatting and structure
- 🎯 Cell range selection for targeted translation
- 🔄 Progress tracking for translation tasks
- 📝 Customizable translation prompts
- 🏢 Field/Industry-specific context support
- 👥 Comparison mode to show original text alongside translations
- ⚙️ User-friendly settings management
- 🖥️ Cross-platform support (Windows, macOS, Linux)

## Requirements

- Python 3.8+
- OpenAI API key

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/excel-gpt-translator.git
cd excel-gpt-translator
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python src/main.py
```

## Usage

1. Launch the application
2. Configure your OpenAI API key in Settings
3. Create a new translation task:
   - Select an Excel file
   - Choose the sheet to translate
   - Specify the cell range (e.g., A1:B4)
   - Select source and target languages
   - Optionally specify the field/industry for context
   - Customize the translation prompt if needed
4. Start the translation task
5. Monitor progress in the task list
6. Access translated files in the same directory as the source file

## Project Structure

```
excel-gpt-translator/
├── src/                    # Source code
│   ├── core/              # Core functionality
│   │   ├── translator.py  # Translation logic
│   │   └── config.py      # Configuration management
│   ├── gui/               # GUI components
│   │   ├── dialogs/       # Dialog windows
│   │   │   ├── task_dialog.py    # Task creation/editing
│   │   │   └── settings_dialog.py # Settings management
│   │   ├── widgets/       # Custom widgets
│   │   │   └── task_widget.py     # Task list item
│   │   └── main_window.py # Main application window
│   └── main.py            # Application entry point
├── tests/                 # Test files
├── requirements.txt       # Python dependencies
└── README.md             # Documentation
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.

The Apache License 2.0 is a permissive open-source license that allows you to:
- Use the software for any purpose
- Modify and distribute the software
- Use the software commercially
- Patent use
- Place warranty

The license requires you to:
- Include the original copyright notice
- Include the Apache License 2.0
- State significant changes made to the software
- Include a NOTICE file if one exists in the original software 