# Excel Formula to Python Function Converter

This project provides tools to automatically convert Excel formulas to Python functions that can be applied to pandas DataFrames.

## Overview

The system consists of three main components:

1. **`prompt_generator.py`** - Analyzes Excel files and generates LLM prompts
2. **`test_formula_conversion.py`** - Pytest unit tests for validation
3. **`sample_data.py`** - Creates sample Excel files for testing

## Installation

```bash
pip install openpyxl pytest pandas
```

## Usage

### Step 1: Generate Prompt for LLM

```bash
python src/prompt_generator.py your_excel_file.xlsx -o generated_prompt.txt
```

This will:
- Create mapping between Excel column letters (A, B, C...) and header names
- Extract unique formula patterns from the Excel file
- Extract data from non-formula columns
- Generate a structured prompt requesting conversion to Python functions

### Step 2: Submit Prompt to LLM

Take the generated prompt from `generated_prompt.txt` and submit it to your preferred LLM (ChatGPT, Claude, etc.). The LLM will generate a Python script that implements the Excel formulas.

### Step 3: Test the Generated Script

```bash
python -m pytest tests/test_formula_conversion.py -v
```

Or test with a specific generated script:

```python
from test_formula_conversion import run_full_test_cycle

success = run_full_test_cycle('your_excel_file.xlsx', 'generated_script.py')
```

## Features

### Prompt Generator Features
- **Smart Formula Pattern Recognition**: Avoids duplicate formulas that differ only in row numbers
- **Header Mapping**: Converts Excel column references (A1, B2) to readable column names
- **Data Schema Extraction**: Provides CSV sample of input data
- **Structured Prompts**: Generates well-formatted requests for LLMs

### Test Suite Features  
- **Automatic Data Extraction**: Creates expected results from original Excel file
- **Script Execution**: Runs generated Python scripts automatically
- **Precision Comparison**: Compares calculated values with floating-point tolerance
- **Comprehensive Validation**: Tests prompt generation, data extraction, and formula conversion

## Example Workflow

1. **Start with Excel file** containing formulas:
   ```
   | Name      | Quantity | Price | Total    | Discount% | Final_Price     |
   |-----------|----------|-------|----------|-----------|-----------------|
   | Product A | 10       | 25.50 | =B2*C2   | 0.1       | =D2*(1-E2)     |
   ```

2. **Generate prompt**:
   ```bash
   python src/prompt_generator.py sample_by_agent/sample_data.xlsx
   ```

3. **Get clean output**:
   ```
   ### Formula 1: Calculate Total
   - **Target Column**: `Total`
   - **Original Excel Formula**: `=B2*C2`
   - **With Column Names**: `=Quantity*Price`
   ```

4. **Submit to LLM** and receive Python script

5. **Test the result**:
```python
from test_formula_conversion import run_full_test_cycle

success = run_full_test_cycle('your_excel_file.xlsx', 'generated_script.py')
```

## Sample Data

Run the sample data generator to create test files:

```bash
python sample_by_agent/sample_data.py
```

This creates `sample_data.xlsx` with:
- Product data (Name, Quantity, Price)
- Formula columns (Total, Final_Price, With_Tax)
- Discount and tax rate columns

## Advanced Usage

### Custom Test Validation

```python
from test_formula_conversion import run_full_test_cycle

# Test your specific files
success = run_full_test_cycle(
    excel_file_path='path/to/your/file.xlsx',
    generated_script_path='path/to/generated/script.py'
)

if success:
    print("✓ All formulas converted correctly!")
else:
    print("✗ Some formulas need adjustment")
```

### Supported Formula Types

Currently supports:
- Basic arithmetic operations (+, -, *, /)
- Parentheses grouping
- Cell references (A1, B2, etc.)
- Common Excel functions (can be extended)

## Files Description

- **`prompt_generator.py`** - Main script for generating LLM prompts
- **`test_formula_conversion.py`** - Comprehensive test suite
- **`sample_data.py`** - Creates sample Excel files for testing
- **`sample_data.xlsx`** - Sample Excel file with formulas
- **`generated_prompt.txt`** - Example generated prompt
- **`improved_prompt.txt`** - Cleaner version of generated prompt

## Testing

The test suite includes:

1. **`test_prompt_generation`** - Validates prompt creation
2. **`test_sample_generated_script`** - Tests with sample LLM output
3. **`test_excel_data_extraction`** - Validates data extraction
4. **`test_formula_identification`** - Tests formula pattern recognition

Run individual tests:
```bash
python -m pytest tests/test_formula_conversion.py::TestFormulaConversion::test_prompt_generation -v
```

## Limitations

- Requires Excel files to have headers in the first row
- Formulas must use standard Excel syntax
- Complex nested functions may need manual adjustment
- Only supports single worksheet analysis

## Contributing

To extend formula support:
1. Add pattern recognition in `_normalize_formula_pattern()`
2. Update `_convert_formula_to_python()` for new operators
3. Add test cases in `test_formula_conversion.py`

## License

This project is provided as-is for educational and development purposes. 