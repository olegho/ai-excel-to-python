#!/usr/bin/env python3
"""
Excel Formula to Python Function Converter

This script converts Excel formulas to Python functions for pandas DataFrames.
Based on the request from generated_prompt.txt file.

Formula to convert:
- unit_price = price/((height*width)/100)
"""

import pandas as pd
import numpy as np
from pathlib import Path


def calculate_unit_price(row):
    """
    Calculate unit_price using the Excel formula: =price/((height*width)/100)
    
    This formula calculates the price per 100 square units of area.
    
    Args:
        row: pandas Series containing the row data
        
    Returns:
        float: The calculated unit price
    """
    try:
        # Calculate area in square units
        area = row['height'] * row['width']
        
        # Convert to price per 100 square units
        area_per_100 = area / 100
        
        # Calculate unit price
        unit_price = row['price'] / area_per_100
        
        return unit_price
    except (ValueError, ZeroDivisionError) as e:
        print(f"Error calculating unit_price for row: {e}")
        return np.nan


def main():
    """Main execution function."""
    
    # Define file paths
    input_file = 'input_data.csv'
    output_file = 'output_data.csv'
    
    try:
        # Read input data
        print(f"Reading data from {input_file}...")
        df = pd.read_csv(input_file)
        
        print(f"Loaded {len(df)} rows with columns: {list(df.columns)}")
        print("\nFirst 5 rows of input data:")
        print(df.head())
        
        # Validate required columns exist
        required_columns = ['width', 'height', 'price']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Missing required columns: {missing_columns}")
        
        # Apply formula to create calculated column
        print("\nApplying formula to calculate unit_price...")
        df['unit_price'] = df.apply(calculate_unit_price, axis=1)
        
        # Display results
        print("\nFirst 5 rows with calculated unit_price:")
        print(df.head())
        
        # Show summary statistics
        print(f"\nSummary statistics for unit_price:")
        print(df['unit_price'].describe())
        
        # Save result
        print(f"\nSaving results to {output_file}...")
        df.to_csv(output_file, index=False)
        
        print("Conversion completed successfully!")
        print(f"Output saved to: {Path(output_file).absolute()}")
        
    except FileNotFoundError:
        print(f"Error: Input file '{input_file}' not found.")
        print("Please ensure the input CSV file exists with columns: width, height, price")
        
        # Create sample data file
        create_sample_data()
        
    except Exception as e:
        print(f"Error during processing: {e}")
        return 1
    
    return 0


def create_sample_data():
    """Create sample input data based on the schema from generated_prompt.txt."""
    
    print("\nCreating sample input data...")
    
    # Sample data from the prompt
    sample_data = {
        'width': [50, 200, 50, 400],
        'height': [100, 80, 50, 200],
        'price': [120, 100, 10, 200]
    }
    
    # Create DataFrame
    df = pd.DataFrame(sample_data)
    
    # Save to CSV
    df.to_csv('input_data.csv', index=False)
    
    print("Sample data created in 'input_data.csv':")
    print(df)
    print("\nYou can now run the script again to process this data.")


if __name__ == "__main__":
    exit_code = main()
    exit(exit_code) 