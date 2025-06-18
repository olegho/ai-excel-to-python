# Excel Formula to Python Function Conversion Request

## Objective
Convert Excel formulas to Python functions that can be applied to pandas DataFrames using the `pd.DataFrame.apply()` method.

## Data Schema and Sample
The input CSV file contains the following columns with sample data (first 10 rows):

```csv
Name,Quantity,Price,Discount%,Tax_Rate
Product A,10,25.5,0.1,0.085
Product B,5,15.75,0.1,0.085
Product C,8,30.0,0.1,0.085
Product D,12,8.25,0.1,0.085
Product E,6,45.0,0.1,0.085
```

## Formulas to Convert
The following formulas need to be converted to Python functions:

### Formula 1
- **Original Excel Formula**: `=B6*C6`
- **With Column Names**: `=Quantity*Price`

### Formula 2
- **Original Excel Formula**: `=B2*C2`
- **With Column Names**: `=Quantity*Price`

### Formula 3
- **Original Excel Formula**: `=F3*(1+G3)`
- **With Column Names**: `=Final_Price*(1+Tax_Rate)`

### Formula 4
- **Original Excel Formula**: `=D4*(1-E4)`
- **With Column Names**: `=Total*(1-Discount%)`

### Formula 5
- **Original Excel Formula**: `=D5*(1-E5)`
- **With Column Names**: `=Total*(1-Discount%)`

### Formula 6
- **Original Excel Formula**: `=B5*C5`
- **With Column Names**: `=Quantity*Price`

### Formula 7
- **Original Excel Formula**: `=F4*(1+G4)`
- **With Column Names**: `=Final_Price*(1+Tax_Rate)`

### Formula 8
- **Original Excel Formula**: `=B3*C3`
- **With Column Names**: `=Quantity*Price`

### Formula 9
- **Original Excel Formula**: `=D2*(1-E2)`
- **With Column Names**: `=Total*(1-Discount%)`

### Formula 10
- **Original Excel Formula**: `=B4*C4`
- **With Column Names**: `=Quantity*Price`

### Formula 11
- **Original Excel Formula**: `=F5*(1+G5)`
- **With Column Names**: `=Final_Price*(1+Tax_Rate)`

### Formula 12
- **Original Excel Formula**: `=F6*(1+G6)`
- **With Column Names**: `=Final_Price*(1+Tax_Rate)`

### Formula 13
- **Original Excel Formula**: `=D3*(1-E3)`
- **With Column Names**: `=Total*(1-Discount%)`

### Formula 14
- **Original Excel Formula**: `=F2*(1+G2)`
- **With Column Names**: `=Final_Price*(1+Tax_Rate)`

### Formula 15
- **Original Excel Formula**: `=D6*(1-E6)`
- **With Column Names**: `=Total*(1-Discount%)`

## Required Output
Please generate a Python script that:

1. **Reads the CSV file** into a pandas DataFrame with the schema shown above
2. **Defines Python functions** for each of the 15 formulas listed above
3. **Creates new columns** using `pd.DataFrame.apply()` to apply these functions
4. **Saves the result** as a new CSV file

### Script Requirements:
- Use `pandas` for data manipulation
- Each formula should be implemented as a separate Python function
- Use `df.apply()` with `axis=1` to apply functions row-wise
- Include proper error handling
- Add comments explaining each function
- The output CSV should contain both original and calculated columns

### Example Structure:
```python
import pandas as pd

def formula_1_function(row):
    # Convert Formula 1: =Quantity*Price
    return # your calculation here

def formula_2_function(row):
    # Convert Formula 2: etc.
    return # your calculation here

# Main execution
df = pd.read_csv('input_data.csv')

# Apply formulas
df['calculated_column_1'] = df.apply(formula_1_function, axis=1)
df['calculated_column_2'] = df.apply(formula_2_function, axis=1)

# Save result
df.to_csv('output_data.csv', index=False)
```

Please provide the complete, runnable Python script.
