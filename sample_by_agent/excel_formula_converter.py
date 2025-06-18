#!/usr/bin/env python3
"""
Excel Formula to Python Function Converter Script

This script reads a CSV file and applies the converted Excel formulas
as Python functions using pandas DataFrame.apply() method.

Generated in response to Excel Formula to Python Function Conversion Request.
"""

import pandas as pd
import sys
from pathlib import Path


def calculate_total(row):
    """
    Calculate Total: =Quantity*Price
    
    Args:
        row: pandas Series representing a single row of data
        
    Returns:
        float: The calculated total (Quantity * Price)
    """
    try:
        return row['Quantity'] * row['Price']
    except (KeyError, TypeError, ValueError) as e:
        print(f"Error calculating Total for row: {e}")
        return 0


def calculate_final_price(row):
    """
    Calculate Final_Price: =Total*(1-Discount%)
    
    Args:
        row: pandas Series representing a single row of data
        
    Returns:
        float: The calculated final price after discount
    """
    try:
        return row['Total'] * (1 - row['Discount%'])
    except (KeyError, TypeError, ValueError) as e:
        print(f"Error calculating Final_Price for row: {e}")
        return 0


def calculate_with_tax(row):
    """
    Calculate With_Tax: =Final_Price*(1+Tax_Rate)
    
    Args:
        row: pandas Series representing a single row of data
        
    Returns:
        float: The calculated price including tax
    """
    try:
        return row['Final_Price'] * (1 + row['Tax_Rate'])
    except (KeyError, TypeError, ValueError) as e:
        print(f"Error calculating With_Tax for row: {e}")
        return 0


def validate_input_data(df):
    """
    Validate that the input DataFrame has all required columns.
    
    Args:
        df: pandas DataFrame to validate
        
    Returns:
        bool: True if valid, False otherwise
    """
    required_columns = ['Name', 'Quantity', 'Price', 'Discount%', 'Tax_Rate']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        print(f"Error: Missing required columns: {missing_columns}")
        print(f"Available columns: {list(df.columns)}")
        return False
    
    # Check for non-numeric values in numeric columns
    numeric_columns = ['Quantity', 'Price', 'Discount%', 'Tax_Rate']
    for col in numeric_columns:
        if not pd.api.types.is_numeric_dtype(df[col]):
            try:
                df[col] = pd.to_numeric(df[col], errors='coerce')
                if df[col].isna().any():
                    print(f"Warning: Non-numeric values found in column '{col}', converted to NaN")
            except Exception as e:
                print(f"Error converting column '{col}' to numeric: {e}")
                return False
    
    return True


def main():
    """
    Main execution function that processes the CSV file and applies formulas.
    """
    # Define input and output file paths
    input_file = 'input_data.csv'
    output_file = 'output_data.csv'
    
    # Check if input file exists
    if not Path(input_file).exists():
        print(f"Error: Input file '{input_file}' not found.")
        print("Please ensure the input CSV file is in the current directory.")
        sys.exit(1)
    
    try:
        # Read input data
        print(f"Reading input data from '{input_file}'...")
        df = pd.read_csv(input_file)
        print(f"Successfully loaded {len(df)} rows of data.")
        
        # Display basic info about the dataset
        print(f"\nDataset Info:")
        print(f"Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        
        # Validate input data
        if not validate_input_data(df):
            print("Input data validation failed.")
            sys.exit(1)
        
        print("\nApplying formulas...")
        
        # Apply Formula 1: Calculate Total
        print("Calculating Total = Quantity * Price...")
        df['Total'] = df.apply(calculate_total, axis=1)
        
        # Apply Formula 2: Calculate Final_Price  
        print("Calculating Final_Price = Total * (1 - Discount%)...")
        df['Final_Price'] = df.apply(calculate_final_price, axis=1)
        
        # Apply Formula 3: Calculate With_Tax
        print("Calculating With_Tax = Final_Price * (1 + Tax_Rate)...")
        df['With_Tax'] = df.apply(calculate_with_tax, axis=1)
        
        # Display sample results
        print(f"\nSample results (first 5 rows):")
        print(df.head().to_string(index=False))
        
        # Save result
        print(f"\nSaving results to '{output_file}'...")
        df.to_csv(output_file, index=False)
        
        # Display summary statistics
        print(f"\nSummary Statistics:")
        print(f"Total rows processed: {len(df)}")
        print(f"Average Total: ${df['Total'].mean():.2f}")
        print(f"Average Final Price: ${df['Final_Price'].mean():.2f}")
        print(f"Average With Tax: ${df['With_Tax'].mean():.2f}")
        
        print("Conversion completed successfully!")
        
    except FileNotFoundError:
        print(f"Error: Could not find the input file '{input_file}'")
        sys.exit(1)
    except pd.errors.EmptyDataError:
        print(f"Error: The input file '{input_file}' is empty")
        sys.exit(1)
    except pd.errors.ParserError as e:
        print(f"Error: Could not parse the CSV file '{input_file}': {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main() 