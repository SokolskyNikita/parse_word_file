# Word Document Processor

A Python script that processes Microsoft Word documents (.docx) using OpenAI's GPT models to either correct grammar or translate text while preserving document formatting.

## Features

- Grammar correction mode to improve English text quality
- Translation mode supporting any target language
- Preserves paragraph formatting including:
  - Basic text styles (bold, italic)
  - Font properties (size, color, caps, etc.)
  - Special text formatting (subscript, superscript, highlighting)
- Skips paragraphs that are too short to process
- Built-in retry, delay and exponential backoff to avoid rate limiting by OpenAI
- Progress tracking with status bar

## Prerequisites

- Python 3.12+
- OpenAI API key

## Installation

1. Clone this repository
2. Install dependencies:
```bash
pip install -r requirements.txt
# Or if you use uv
uv pip install --system -r requirements.txt
```
3. Set your OpenAI API key:
```bash
export OPENAI_API_KEY='your-api-key-here'
```

## Usage

Grammar correction (default mode):
```bash
python process_word_file.py document.docx
# or explicitly
python process_word_file.py document.docx -fg
```

Translation:
```bash
python process_word_file.py document.docx -t French
python process_word_file.py document.docx -t Russian
```

Custom output path:
```bash
python process_word_file.py document.docx -t Spanish -o translated.docx
```

## Arguments

- `input`: Path to the input DOCX file
- `-fg, --fix-grammar`: Use grammar correction mode (default)
- `-t LANG, --translate LANG`: Translate to specified language
- `-o PATH, --output PATH`: Custom output file path
- `--model MODEL`: OpenAI model to use (default: gpt-4o)

## Output

The script creates a new Word document with processed text while maintaining the original formatting. Default output filenames are:
- Grammar mode: `<input>_grammar_fixed.docx`
- Translation mode: `<input>_translated_to_<language>.docx` 