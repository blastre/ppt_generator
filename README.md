# CSV to PowerPoint Generator

Convert CSV data into professional PowerPoint presentations with AI-powered insights and interactive chatbot functionality.

## Features

- **AI-Powered Analysis**: Automatically generates insights from CSV data using Groq LLM
- **Professional PPT Generation**: Creates polished presentations with charts and bullet points
- **Interactive Chatbot**: Ask questions about your data in real-time
- **Multiple Templates**: Choose from 6+ professional PowerPoint templates
- **Smart Chart Generation**: Automatic bar and pie chart creation with Plotly/matplotlib
- **CLI Interface**: Easy command-line usage
- **Python Integration**: Import functions for seamless integration into other projects

## Quick Start

### CLI Usage

```bash
# Generate PowerPoint from CSV
python main.py data.csv "What are the sales trends?" -t modern_blue

# Interactive chatbot mode
python main.py data.csv --chat

# List available templates
python main.py --list-templates
```

### Python Integration

```python
from main import csv_to_ppt, run_chatbot
from chatbot import chatbot

# Generate presentation
ppt_path = csv_to_ppt("data.csv", "Analyze revenue by region", template="corporate_green")

# Start chatbot session
run_chatbot("data.csv")

# Single question answering
from csv_processor import load_csv, analyze_data
df = load_csv("data.csv")
result = analyze_data(df, "What's the average sales?")
answer = chatbot(result['summary'], "How can we improve performance?")
```

## Installation

```bash
pip install pandas matplotlib plotly python-pptx groq
```

Set your Groq API key in `model.py`:
```python
GROQ_API_KEY = "your_api_key_here"
```

## Templates Available

- `default` - Professional blue template
- `modern_blue` - Clean modern design
- `corporate_green` - Business-focused green theme
- `minimalist_gray` - Simple gray with blue accents
- `vibrant_orange` - Creative orange theme
- `elegant_purple` - Sophisticated purple styling

## Project Structure

- `main.py` - CLI interface and main functions
- `csv_processor.py` - Data analysis and pandas operations
- `ppt_generator.py` - PowerPoint creation with templates
- `chart_generator.py` - Plotly/matplotlib chart generation
- `chatbot.py` - Interactive Q&A functionality
- `model.py` - AI model integration (Groq)

## Key Functions for Integration

- `csv_to_ppt(csv_file, question, output_file, template)` - Generate PPT
- `run_chatbot(csv_file)` - Start interactive session
- `chatbot(summary, question)` - Single Q&A
- `analyze_data(df, question)` - Data analysis
- `load_csv(filepath)` - CSV loading with pandas

## Example Output

Each presentation includes:
1. Executive Summary slide
2. Data Overview & Methodology
3. Primary insights with charts
4. Detailed analysis with visualizations
5. Strategic recommendations

## License

MIT License - Feel free to integrate into your projects.
