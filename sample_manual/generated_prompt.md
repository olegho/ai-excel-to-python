# Excel Formula to Python Function Conversion Request

## Objective
Convert Excel formulas to Python functions that can be applied to pandas DataFrames using the `pd.DataFrame.apply()` method.

## Data Schema and Sample
The input CSV file contains the following columns with sample data (first 10 rows):

```csv
width,height,price
50,100,120
200,80,100
50,50,10
400,200,200
```

## Formulas to Convert
The following formulas need to be converted to Python functions:

### Formula 1: Calculate unit_price
- **Target Column**: `unit_price`
- **Original Excel Formula**: `=C2/((B2*A2)/100)`
- **With Column Names**: `=price/((height*width)/100)`

## Required Output
Please generate a Python script that:

1. **Reads the CSV file** into a pandas DataFrame with the schema shown above
2. **Defines Python functions** for each of the 1 formulas listed above
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

def calculate_unit_price(row):
    '''Calculate unit_price: =price/((height*width)/100)'''
    return row['price'] / ((row['height'] * row['width']) / 100)

def calculate_second_formula(row):
    '''Calculate Second Formula: etc.'''
    return # your calculation here

# Main execution
if __name__ == "__main__":
    # Read input data
    df = pd.read_csv('input_data.csv')
    
    # Apply formulas to create calculated columns
    df['unit_price'] = df.apply(calculate_unit_price, axis=1)
    df['Second_Column'] = df.apply(calculate_second_formula, axis=1)
    
    # Save result
    df.to_csv('output_data.csv', index=False)
    print("Conversion completed successfully!")
```

Please provide the complete, runnable Python script that implements all the formulas shown above.
