# Word Document Grammar Parser

A Python script that processes Microsoft Word documents (.docx) using OpenAI's GPT models to correct grammar or translate text while preserving document formatting.

## Features

- Processes .docx files while maintaining basic text formatting
- Grammar correction mode to improve text quality
- Translation mode supporting any target language
- Preserves essential formatting including bold, italic, font properties, and more
- Includes retry logic for API calls
- Progress bar for tracking document processing
- Configurable model selection

## Prerequisites

- Python 3.12+
- OpenAI API key

## Installation

1. Clone this repository
2. Install dependencies:
```bash
pip install -r requirements.txt
```
3. Set your OpenAI API key as an environment variable:
```bash
export OPENAI_API_KEY='your-api-key-here'
```

## Usage

Grammar correction:
```bash
python process_word_file.py input.docx --fix-grammar
```

Translation:
```bash
python process_word_file.py input.docx --translate-to "French"
```

With custom output path:
```bash
python process_word_file.py input.docx --translate-to "Spanish" -o output.docx
```

With specific model:
```bash
python process_word_file.py input.docx --fix-grammar --model gpt-4
```

## Arguments

- `input`: Path to the input DOCX file
- `-o, --output`: Output DOCX file path (default: <input>_grammar_fixed.docx or <input>_translated_to_<lang>.docx)
- `--model`: OpenAI model to use (default: gpt-4o)
- `--fix-grammar`: Enable grammar correction mode
- `--translate-to`: Translate the text to the specified language (e.g., "French", "Spanish", "Japanese")

Note: You must specify either `--fix-grammar` or `--translate-to`, but not both.

## Output

The script creates a new Word document with processed text while maintaining the basic formatting of the original document. The output filename will depend on the processing mode:
- For grammar correction: `<input>_grammar_fixed.docx`
- For translation: `<input>_translated_to_<language>.docx` 