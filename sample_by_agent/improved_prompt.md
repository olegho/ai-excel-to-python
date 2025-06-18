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

### Formula 1: Calculate Total
- **Target Column**: `Total`
- **Original Excel Formula**: `=B2*C2`
- **With Column Names**: `=Quantity*Price`

### Formula 2: Calculate Final_Price
- **Target Column**: `Final_Price`
- **Original Excel Formula**: `=D2*(1-E2)`
- **With Column Names**: `=Total*(1-Discount%)`

### Formula 3: Calculate With_Tax
- **Target Column**: `With_Tax`
- **Original Excel Formula**: `=F2*(1+G2)`
- **With Column Names**: `=Final_Price*(1+Tax_Rate)`

## Required Output
Please generate a Python script that:

1. **Reads the CSV file** into a pandas DataFrame with the schema shown above
2. **Defines Python functions** for each of the 3 formulas listed above
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

def calculate_total(row):
    '''Calculate Total: =Quantity*Price'''
    return row['Quantity'] * row['Price']

def calculate_final_price(row):
    '''Calculate Final_Price: =Total*(1-Discount%)'''
    return # your calculation here

# Main execution
if __name__ == "__main__":
    # Read input data
    df = pd.read_csv('input_data.csv')
    
    # Apply formulas to create calculated columns
    df['Total'] = df.apply(calculate_total, axis=1)
    df['Final_Price'] = df.apply(calculate_final_price, axis=1)
    
    # Save result
    df.to_csv('output_data.csv', index=False)
    print("Conversion completed successfully!")
```

Please provide the complete, runnable Python script that implements all the formulas shown above.
