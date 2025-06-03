# Enhanced Excel Question Extraction & Filling System

A sophisticated Python system for intelligently extracting questions from Excel documents and filling them with contextually appropriate responses.

## 🚀 Features

- **🔍 Intelligent Content Sheet Detection**: Automatically identifies instruction/content sheets without hardcoding
- **📊 Statistical Column Analysis**: Uses data patterns to identify question and answer columns
- **🧠 Smart Hierarchy Parsing**: LLM-powered understanding of document structure and question relationships
- **✏️ Cross-Column Intelligent Filling**: Applies business logic across multiple columns
- **📚 Global Context Awareness**: Extracts document-wide context to guide processing
- **🎯 Multiple Processing Modes**: Extract-only, fill-only, or combined processing

## 📁 Project Structure

```
enhanced-excel-system/
├── main.py              # Main entry point
├── models.py            # Pydantic data models
├── claude_client.py     # Claude API client
├── parsers.py           # Question parsing utilities
├── extractor.py         # Excel extraction logic
├── filler.py            # Excel filling logic
├── utils.py             # Utility functions
├── config.py            # Configuration settings
├── requirements.txt     # Python dependencies
├── .env.example         # Environment variables template
└── README.md           # This file
```

## 🛠️ Installation

1. **Clone or download the project files**

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up environment variables:**
   ```bash
   cp .env.example .env
   # Edit .env and add your Anthropic API key
   ```

4. **Get an Anthropic API key:**
   - Visit [Anthropic Console](https://console.anthropic.com/)
   - Create an account and generate an API key
   - Add it to your `.env` file

## 📖 Usage

### Basic Commands

```bash
# Extract questions only
python main.py extract document.xlsx

# Fill existing extraction
python main.py fill document.xlsx extraction_results.json

# Extract and fill (complete process)
python main.py both document.xlsx
```

### Advanced Options

```bash
# Use rule-based hierarchy parsing (faster, less accurate)
python main.py extract document.xlsx --no-llm-hierarchy

# Manually specify content sheet
python main.py extract document.xlsx --content-sheet="Instructions"

# Complete process with options
python main.py both document.xlsx --content-sheet="Guidelines" --no-llm-hierarchy
```

## 🔧 How It Works

### 1. Intelligent Content Detection
- Automatically scans all sheets to identify instruction/content sheets
- Uses pattern recognition to distinguish between content and question sheets
- Extracts global context and filling guidelines

### 2. Statistical Column Analysis
- Analyzes text length patterns, fill ratios, and content types
- Identifies question columns (longer descriptive text)
- Detects answer columns (shorter responses, specific headers)

### 3. Smart Question Extraction
- Parses multi-line questions within single cells
- Uses LLM to understand hierarchical relationships
- Classifies questions by type (headers, requirements, sub-items)

### 4. Intelligent Filling
- Generates realistic responses based on document context
- Applies cross-column business logic
- Maintains consistency across related fields

## 📊 Output Files

### Extraction Mode
```
enhanced_extraction_[filename]_[timestamp]/
├── extraction_results.json    # Complete extraction data
└── analysis_report.json      # Statistical analysis and summary
```

### Filling Mode
```
enhanced_filled_[timestamp]_[filename].xlsx
```

## 🎯 Supported Document Types

- **RFI/RFP Documents**: Request for Information/Proposal
- **Compliance Checklists**: Regulatory and standard compliance
- **Technical Questionnaires**: Product and service capabilities
- **Assessment Forms**: Evaluation and scoring documents

## ⚡ Performance Tips

1. **Use `--no-llm-hierarchy` for faster processing** (trades accuracy for speed)
2. **Specify content sheets manually** when known to skip detection
3. **Process smaller files first** to test extraction patterns
4. **Review analysis reports** to understand document structure

## 🔍 Troubleshooting

### Common Issues

**"ANTHROPIC_API_KEY not found"**
- Ensure `.env` file exists with your API key
- Verify the key is valid and has sufficient credits

**"No questions extracted"**
- Check if sheets contain actual questions vs. just data
- Try manually specifying the content sheet
- Review the analysis report for detection issues

**"Column detection failed"**
- Verify the sheet has clear question and answer column structure
- Check if headers are descriptive (e.g., "Compliance", "Response")

### Debug Mode

Add debug prints by modifying `config.py`:
```python
DEBUG_MODE = True
```

## 🤝 Contributing

This system is designed to be modular and extensible:

- **Add new document types**: Extend pattern recognition in `claude_client.py`
- **Improve parsing**: Enhance hierarchy detection in `parsers.py`  
- **Add filling strategies**: Extend cross-column logic in `filler.py`
- **New output formats**: Add exporters in `utils.py`

## 📄 License

This project is provided as-is for educational and commercial use. Please ensure compliance with your organization's policies regarding AI-assisted document processing.

## 🆘 Support

For issues or questions:
1. Check the troubleshooting section above
2. Review the analysis reports for insight into processing
3. Test with simpler documents first to isolate issues
4. Ensure your Excel files have clear question/answer column structure

---

**Built with intelligence, designed for efficiency** 🚀