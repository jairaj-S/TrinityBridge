#!/usr/bin/env python3
"""
Raw Material Supplier Analysis Tool
Evaluates suppliers based on price, tariffs, and maturity ratings to support procurement decisions.
"""

import argparse
import logging
import sys
import os
from typing import Optional, List, Dict, Any, Union
from pathlib import Path

import pandas as pd

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# Tariff rates mapping (country_code, raw_material) to tariff percentage
TARIFF_RATES = {
    ('CN', 'Steel'): 0.25,          # 25% tariff on Chinese steel
    ('US', 'Steel'): 0.05,          # 5% tariff on US steel
    ('DE', 'Steel'): 0.07,          # 7% tariff on German steel
    ('IN', 'Steel'): 0.20,          # 20% tariff on Indian steel
    ('JP', 'Steel'): 0.08,          # 8% tariff on Japanese steel
    ('CN', 'Aluminum'): 0.15,       # 15% tariff on Chinese aluminum
    ('US', 'Aluminum'): 0.03,       # 3% tariff on US aluminum
    ('DE', 'Aluminum'): 0.05,       # 5% tariff on German aluminum
}

def get_tariff_rate(country_of_origin: str, raw_material: str) -> float:
    """
    Get the tariff rate for a specific country and material combination.
    
    Args:
        country_of_origin: 2-letter country code
        raw_material: Name of the raw material
        
    Returns:
        Tariff rate as a decimal (e.g., 0.10 for 10%)
    """
    key = (country_of_origin.upper(), raw_material)
    return TARIFF_RATES.get(key, 0.0)  # Default to 0% if not found

def calculate_tariffed_price(base_price: float, country_of_origin: str, raw_material: str) -> float:
    """
    Calculate the price after applying relevant tariff.
    
    Args:
        base_price: Base price of the raw material
        country_of_origin: 2-letter country code
        raw_material: Name of the raw material
        
    Returns:
        Price after applying tariff (base_price * (1 + tariff_rate))
    """
    tariff_rate = get_tariff_rate(country_of_origin, raw_material)
    return base_price * (1 + tariff_rate)

def calculate_price_score(prices: List[float]) -> List[float]:
    """
    Calculate price scores using inverse scoring where lower prices get higher scores.
    Normalizes scores so the highest score is 100.
    
    Args:
        prices: List of prices (after tariffs) from different suppliers
        
    Returns:
        List of normalized price scores (0-100), where 100 is the best (lowest) price
        
    Raises:
        ValueError: If prices list is empty or contains non-positive values
    """
    if not prices:
        raise ValueError("Prices list cannot be empty")
    
    if any(price <= 0 for price in prices):
        raise ValueError("All prices must be positive values")
    
    # Calculate inverse scores (1/price)
    inverse_scores = [1/price for price in prices]
    
    # Normalize to 0-100 scale
    max_inverse = max(inverse_scores)
    normalized_scores = [(score / max_inverse) * 100 for score in inverse_scores]
    
    return normalized_scores

def convert_maturity_to_score(maturity_input: Union[str, float, int, None]) -> float:
    """
    Convert maturity rating to numerical score.
    
    Args:
        maturity_input: Either qualitative string ('High', 'Medium', 'Low') 
                       or quantitative number (0.0-1.0), or None
        
    Returns:
        Numerical maturity score (0-100)
        
    Raises:
        ValueError: If input is invalid qualitative rating or out of range quantitative value
    """
    if maturity_input is None:
        return 50.0  # Default neutral score
    
    # Handle string input (qualitative)
    if isinstance(maturity_input, str):
        qualitative_mapping = {
            'high': 100.0,
            'medium': 50.0,
            'low': 20.0
        }
        key = maturity_input.strip().lower()
        if key not in qualitative_mapping:
            raise ValueError(f"Invalid qualitative maturity rating: {maturity_input}. Must be 'High', 'Medium', or 'Low'")
        return qualitative_mapping[key]
    
    # Handle numeric input (quantitative)
    if isinstance(maturity_input, (int, float)):
        value = float(maturity_input)
        
        # Handle 0-1 scale
        if 0.0 <= value <= 1.0:
            return value * 100.0
        
        # Handle 0-100 scale
        if 0.0 <= value <= 100.0:
            return value
        
        raise ValueError(f"Quantitative maturity rating must be between 0.0-1.0 or 0.0-100.0, got: {value}")
    
    raise ValueError(f"Maturity input must be string, number, or None, got: {type(maturity_input)}")

def calculate_reliability_score(price_score: float, maturity_score: float, 
                              price_weight: float = 0.6, maturity_weight: float = 0.4) -> float:
    """
    Calculate the final reliability score using weighted combination of price and maturity scores.
    
    Args:
        price_score: Price score (0-100)
        maturity_score: Maturity score (0-100)
        price_weight: Weight for price score (default: 0.6)
        maturity_weight: Weight for maturity score (default: 0.4)
        
    Returns:
        Weighted reliability score (0-100)
        
    Raises:
        ValueError: If weights don't sum to 1.0 or scores are out of range
    """
    # Validate weights
    if abs(price_weight + maturity_weight - 1.0) > 0.001:
        raise ValueError(f"Weights must sum to 1.0, got: {price_weight + maturity_weight}")
    
    # Validate score ranges
    if not (0 <= price_score <= 100):
        raise ValueError(f"Price score must be between 0-100, got: {price_score}")
    if not (0 <= maturity_score <= 100):
        raise ValueError(f"Maturity score must be between 0-100, got: {maturity_score}")
    
    # Calculate weighted score
    reliability_score = (price_score * price_weight) + (maturity_score * maturity_weight)
    
    return reliability_score

def read_excel_input(excel_file: str) -> Dict[str, Any]:
    """
    Read supplier data from Excel file.
    
    Expected Excel format:
    - Sheet name: 'Suppliers' (or first sheet)
    - Columns: name, price_per_rm, country_of_origin, maturity_rating_qualitative, maturity_rating_quantitative
    - Additional sheet metadata: raw_material, quantity_needed (can be in first row or separate sheet)
    
    Args:
        excel_file: Path to Excel file
        
    Returns:
        Dictionary with raw_material, quantity_needed, and suppliers_data
        
    Raises:
        FileNotFoundError: If Excel file doesn't exist
        ValueError: If Excel file format is invalid
    """
    try:
        # Read Excel file
        excel_data = pd.ExcelFile(excel_file)
        
        # Try to read configuration from a 'Config' sheet first
        raw_material = None
        quantity_needed = None
        
        if 'Config' in excel_data.sheet_names:
            config_df = pd.read_excel(excel_file, sheet_name='Config')
            if 'raw_material' in config_df.columns and len(config_df) > 0:
                raw_material = config_df['raw_material'].iloc[0]
            if 'quantity_needed' in config_df.columns and len(config_df) > 0:
                quantity_needed = config_df['quantity_needed'].iloc[0]
        
        # Read suppliers data
        suppliers_sheet = 'Suppliers' if 'Suppliers' in excel_data.sheet_names else excel_data.sheet_names[0]
        suppliers_df = pd.read_excel(excel_file, sheet_name=suppliers_sheet)
        
        # If config not found in separate sheet, try to read from suppliers sheet metadata
        if raw_material is None and 'raw_material' in suppliers_df.columns:
            raw_material = suppliers_df['raw_material'].iloc[0] if len(suppliers_df) > 0 else None
        if quantity_needed is None and 'quantity_needed' in suppliers_df.columns:
            quantity_needed = suppliers_df['quantity_needed'].iloc[0] if len(suppliers_df) > 0 else None
        
        # Validate required columns
        required_columns = ['name', 'price_per_rm', 'country_of_origin']
        missing_columns = [col for col in required_columns if col not in suppliers_df.columns]
        if missing_columns:
            raise ValueError(f"Missing required columns in Excel file: {missing_columns}")
        
        # Convert DataFrame to list of dictionaries
        suppliers_data = []
        for _, row in suppliers_df.iterrows():
            supplier = {
                'name': str(row['name']),
                'price_per_rm': float(row['price_per_rm']),
                'country_of_origin': str(row['country_of_origin']).upper()
            }
            
            # Add optional maturity ratings
            if 'maturity_rating_qualitative' in suppliers_df.columns:
                qual_value = row['maturity_rating_qualitative']
                if pd.notna(qual_value):
                    supplier['maturity_rating_qualitative'] = str(qual_value)
            if 'maturity_rating_quantitative' in suppliers_df.columns:
                quant_value = row['maturity_rating_quantitative']
                if pd.notna(quant_value):
                    supplier['maturity_rating_quantitative'] = float(quant_value)
            
            suppliers_data.append(supplier)
        
        # Use defaults if not specified
        if raw_material is None:
            raw_material = "Unknown Material"
        if quantity_needed is None:
            quantity_needed = 1000  # Default quantity
        
        return {
            'raw_material': raw_material,
            'quantity_needed': float(quantity_needed),
            'suppliers_data': suppliers_data
        }
        
    except FileNotFoundError:
        raise FileNotFoundError(f"Excel file not found: {excel_file}")
    except Exception as e:
        raise ValueError(f"Error reading Excel file: {str(e)}")

def export_results_to_excel(analysis_results: List[Dict[str, Any]], output_file: str, 
                          raw_material: str, quantity_needed: Union[float, int]) -> None:
    """
    Export analysis results to Excel file with multiple sheets.
    
    Args:
        analysis_results: List of analysis result dictionaries
        output_file: Path for output Excel file
        raw_material: Name of the raw material
        quantity_needed: Quantity needed
    """
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Summary sheet with key results
            summary_data = []
            for i, result in enumerate(analysis_results, 1):
                summary_data.append({
                    'Rank': i,
                    'Supplier': result['supplier_name'],
                    'Country': result['details']['country_of_origin'],
                    'Total Cost': result['price_after_tariff_for_needed_quantity'],
                    'Reliability Score': result['reliability_score_after_tariff'],
                    'Base Price/Unit': result['details']['base_price_per_unit'],
                    'Tariffed Price/Unit': result['details']['tariffed_price_per_unit'],
                    'Price Score': result['details']['price_score'],
                    'Maturity Score': result['details']['maturity_score']
                })
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Detailed analysis sheet
            detailed_data = []
            for result in analysis_results:
                details = result['details']
                detailed_data.append({
                    'Raw Material': result['raw_material'],
                    'Supplier Name': result['supplier_name'],
                    'Country of Origin': details['country_of_origin'],
                    'Quantity Needed': quantity_needed,
                    'Base Price per Unit': details['base_price_per_unit'],
                    'Tariffed Price per Unit': details['tariffed_price_per_unit'],
                    'Total Cost (Tariffed)': result['price_after_tariff_for_needed_quantity'],
                    'Price Score (0-100)': details['price_score'],
                    'Maturity Score (0-100)': details['maturity_score'],
                    'Reliability Score (0-100)': result['reliability_score_after_tariff'],
                    'Price Weight Used': details['weights_used']['price_weight'],
                    'Maturity Weight Used': details['weights_used']['maturity_weight']
                })
            
            detailed_df = pd.DataFrame(detailed_data)
            detailed_df.to_excel(writer, sheet_name='Detailed Analysis', index=False)
            
            # Cost comparison sheet
            best_supplier = analysis_results[0]
            most_expensive = max(analysis_results, key=lambda x: x['price_after_tariff_for_needed_quantity'])
            cheapest = min(analysis_results, key=lambda x: x['price_after_tariff_for_needed_quantity'])
            
            comparison_data = [
                {
                    'Metric': 'Best Overall Choice (Highest Reliability)',
                    'Supplier': best_supplier['supplier_name'],
                    'Cost': best_supplier['price_after_tariff_for_needed_quantity'],
                    'Reliability Score': best_supplier['reliability_score_after_tariff']
                },
                {
                    'Metric': 'Cheapest Option',
                    'Supplier': cheapest['supplier_name'],
                    'Cost': cheapest['price_after_tariff_for_needed_quantity'],
                    'Reliability Score': cheapest['reliability_score_after_tariff']
                },
                {
                    'Metric': 'Most Expensive Option',
                    'Supplier': most_expensive['supplier_name'],
                    'Cost': most_expensive['price_after_tariff_for_needed_quantity'],
                    'Reliability Score': most_expensive['reliability_score_after_tariff']
                }
            ]
            
            comparison_df = pd.DataFrame(comparison_data)
            comparison_df.to_excel(writer, sheet_name='Cost Comparison', index=False)
            
            # Analysis metadata
            metadata = pd.DataFrame([
                {'Parameter': 'Raw Material', 'Value': raw_material},
                {'Parameter': 'Quantity Needed', 'Value': quantity_needed},
                {'Parameter': 'Number of Suppliers Evaluated', 'Value': len(analysis_results)},
                {'Parameter': 'Best Supplier', 'Value': best_supplier['supplier_name']},
                {'Parameter': 'Best Supplier Cost', 'Value': best_supplier['price_after_tariff_for_needed_quantity']},
                {'Parameter': 'Best Supplier Reliability Score', 'Value': best_supplier['reliability_score_after_tariff']},
                {'Parameter': 'Cost Range', 'Value': f"${cheapest['price_after_tariff_for_needed_quantity']:,.2f} - ${most_expensive['price_after_tariff_for_needed_quantity']:,.2f}"},
                {'Parameter': 'Potential Savings vs Most Expensive', 'Value': f"${most_expensive['price_after_tariff_for_needed_quantity'] - best_supplier['price_after_tariff_for_needed_quantity']:,.2f}"}
            ])
            metadata.to_excel(writer, sheet_name='Analysis Info', index=False)
        
        print(f"Analysis results exported to: {output_file}")
        
    except Exception as e:
        raise ValueError(f"Error exporting to Excel: {str(e)}")

def process_supplier_data(supplier_data: Dict[str, Any], raw_material: str) -> Dict[str, Any]:
    """
    Process individual supplier data to calculate price after tariff and maturity score.
    
    Args:
        supplier_data: Dictionary containing supplier information
        raw_material: Name of the raw material
        
    Returns:
        Dictionary with processed supplier information
    """
    # Calculate price after tariff
    price_after_tariff = calculate_tariffed_price(
        supplier_data['price_per_rm'],
        supplier_data['country_of_origin'],
        raw_material
    )
    
    # Calculate maturity score
    if supplier_data.get('maturity_rating_qualitative'):
        maturity_score = convert_maturity_to_score(supplier_data['maturity_rating_qualitative'])
    elif supplier_data.get('maturity_rating_quantitative') is not None:
        maturity_score = convert_maturity_to_score(supplier_data['maturity_rating_quantitative'])
    else:
        maturity_score = convert_maturity_to_score(None)
    
    return {
        'supplier_name': supplier_data['name'],
        'price_after_tariff': price_after_tariff,
        'maturity_score': maturity_score,
        'country_of_origin': supplier_data['country_of_origin'],
        'base_price': supplier_data['price_per_rm']
    }

def analyze_raw_material_suppliers(raw_material: str, quantity_needed: Union[float, int], 
                                 suppliers_data: List[Dict[str, Any]], 
                                 price_weight: float = 0.6, maturity_weight: float = 0.4) -> List[Dict[str, Any]]:
    """
    Orchestrate the entire supplier analysis process.
    
    Args:
        raw_material: Name of the raw material
        quantity_needed: Quantity needed
        suppliers_data: List of supplier dictionaries
        price_weight: Weight for price score (default: 0.6)
        maturity_weight: Weight for maturity score (default: 0.4)
        
    Returns:
        List of dictionaries with comprehensive supplier analysis
    """
    # Initialize list to store processed supplier data
    processed_suppliers = []
    
    # Process each supplier
    for supplier_data in suppliers_data:
        processed_supplier = process_supplier_data(supplier_data, raw_material)
        processed_suppliers.append(processed_supplier)
    
    # Collect all tariffed prices for price scoring
    tariffed_prices = [supplier['price_after_tariff'] for supplier in processed_suppliers]
    
    # Calculate price scores using inverse scoring
    price_scores = calculate_price_score(tariffed_prices)
    
    # Build final analysis results
    analysis_results = []
    for i, processed_supplier in enumerate(processed_suppliers):
        price_score = price_scores[i]
        maturity_score = processed_supplier['maturity_score']
        
        # Calculate reliability score
        reliability_score = calculate_reliability_score(
            price_score, maturity_score, price_weight, maturity_weight
        )
        
        # Calculate total cost for needed quantity
        price_after_tariff_for_quantity = processed_supplier['price_after_tariff'] * quantity_needed
        
        analysis_results.append({
            'raw_material': raw_material,
            'supplier_name': processed_supplier['supplier_name'],
            'price_after_tariff_for_needed_quantity': price_after_tariff_for_quantity,
            'reliability_score_after_tariff': reliability_score,
            'details': {
                'country_of_origin': processed_supplier['country_of_origin'],
                'base_price_per_unit': processed_supplier['base_price'],
                'tariffed_price_per_unit': processed_supplier['price_after_tariff'],
                'price_score': price_score,
                'maturity_score': maturity_score,
                'weights_used': {
                    'price_weight': price_weight,
                    'maturity_weight': maturity_weight
                }
            }
        })
    
    # Sort by reliability score (descending)
    analysis_results.sort(key=lambda x: x['reliability_score_after_tariff'], reverse=True)
    
    return analysis_results

def create_sample_excel(output_file: str) -> None:
    """
    Create a sample Excel input file with example supplier data.
    
    Args:
        output_file: Path for the sample Excel file
    """
    try:
        # Create sample data matching the expected format
        sample_suppliers = [
            {
                'name': 'Global Steel Corp',
                'price_per_rm': 45.50,
                'country_of_origin': 'US',
                'maturity_rating_qualitative': 'High',
                'maturity_rating_quantitative': None
            },
            {
                'name': 'European Steel Solutions',
                'price_per_rm': 52.00,
                'country_of_origin': 'DE',
                'maturity_rating_qualitative': 'High',
                'maturity_rating_quantitative': None
            },
            {
                'name': 'Budget Steel Inc',
                'price_per_rm': 29.99,
                'country_of_origin': 'IN',
                'maturity_rating_qualitative': 'Low',
                'maturity_rating_quantitative': None
            },
            {
                'name': 'Premium Metals Ltd',
                'price_per_rm': 41.75,
                'country_of_origin': 'JP',
                'maturity_rating_qualitative': None,
                'maturity_rating_quantitative': 0.5
            },
            {
                'name': 'Asian Metal Works',
                'price_per_rm': 38.20,
                'country_of_origin': 'CN',
                'maturity_rating_qualitative': 'Medium',
                'maturity_rating_quantitative': None
            }
        ]
        
        # Create configuration data
        config_data = [
            {'raw_material': 'Steel', 'quantity_needed': 1000}
        ]
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Write suppliers data
            suppliers_df = pd.DataFrame(sample_suppliers)
            suppliers_df.to_excel(writer, sheet_name='Suppliers', index=False)
            
            # Write configuration
            config_df = pd.DataFrame(config_data)
            config_df.to_excel(writer, sheet_name='Config', index=False)
            
            # Add instructions sheet
            instructions = pd.DataFrame([
                {'Instructions': 'How to use this Excel template:'},
                {'Instructions': ''},
                {'Instructions': '1. Config Sheet:'},
                {'Instructions': '   - raw_material: Name of the material you are analyzing'},
                {'Instructions': '   - quantity_needed: Total quantity required'},
                {'Instructions': ''},
                {'Instructions': '2. Suppliers Sheet:'},
                {'Instructions': '   - name: Supplier company name'},
                {'Instructions': '   - price_per_rm: Price per unit of raw material'},
                {'Instructions': '   - country_of_origin: 2-letter country code (US, CN, DE, etc.)'},
                {'Instructions': '   - maturity_rating_qualitative: High/Medium/Low (optional)'},
                {'Instructions': '   - maturity_rating_quantitative: 0.0-1.0 decimal value (optional)'},
                {'Instructions': ''},
                {'Instructions': '3. Run analysis:'},
                {'Instructions': '   python main.py --input-excel your_file.xlsx --output-excel results.xlsx'},
                {'Instructions': ''},
                {'Instructions': '4. Output will contain multiple sheets:'},
                {'Instructions': '   - Summary: Key results ranked by reliability score'},
                {'Instructions': '   - Detailed Analysis: Complete breakdown of all metrics'},
                {'Instructions': '   - Cost Comparison: Best vs worst options'},
                {'Instructions': '   - Analysis Info: Metadata and parameters used'}
            ])
            instructions.to_excel(writer, sheet_name='Instructions', index=False)
        
        print(f"Sample Excel file created: {output_file}")
        print("You can now modify the data and run analysis with:")
        print(f"python main.py --input-excel {output_file} --output-excel results.xlsx")
        
    except Exception as e:
        raise ValueError(f"Error creating sample Excel file: {str(e)}")

def analyze_from_excel(input_excel: str, output_excel: Optional[str] = None, 
                      price_weight: float = 0.6, maturity_weight: float = 0.4) -> List[Dict[str, Any]]:
    """
    Complete Excel-to-Excel analysis workflow.
    
    Args:
        input_excel: Path to input Excel file with supplier data
        output_excel: Path for output Excel file (optional)
        price_weight: Weight for price score (default: 0.6)
        maturity_weight: Weight for maturity score (default: 0.4)
        
    Returns:
        List of analysis results
    """
    # Read input data from Excel
    input_data = read_excel_input(input_excel)
    
    # Run analysis
    results = analyze_raw_material_suppliers(
        input_data['raw_material'],
        input_data['quantity_needed'],
        input_data['suppliers_data'],
        price_weight,
        maturity_weight
    )
    
    # Export to Excel if output file specified
    if output_excel:
        export_results_to_excel(
            results, 
            output_excel, 
            input_data['raw_material'], 
            input_data['quantity_needed']
        )
    
    return results

def get_sample_data() -> Dict[str, Any]:
    """Get sample input data for testing."""
    return {
        'raw_material': 'Steel',
        'quantity_needed': 1000,
        'suppliers_data': [
            {
                'name': 'Global Steel Corp',
                'price_per_rm': 45.50,
                'country_of_origin': 'US',
                'maturity_rating_qualitative': 'High'
            },
            {
                'name': 'European Steel Solutions',
                'price_per_rm': 52.00,
                'country_of_origin': 'DE',
                'maturity_rating_qualitative': 'High'
            },
            {
                'name': 'Budget Steel Inc',
                'price_per_rm': 29.99,
                'country_of_origin': 'IN',
                'maturity_rating_qualitative': 'Low'
            },
            {
                'name': 'Premium Metals Ltd',
                'price_per_rm': 41.75,
                'country_of_origin': 'JP',
                'maturity_rating_quantitative': 0.5
            },
            {
                'name': 'Asian Metal Works',
                'price_per_rm': 38.20,
                'country_of_origin': 'CN',
                'maturity_rating_qualitative': 'Medium'
            }
        ]
    }

def main_demo():
    """Demonstration of the complete supplier analysis functionality."""
    print("=== RAW MATERIAL SUPPLIER ANALYSIS DEMO ===\n")
    
    # Get sample input data
    sample_data = get_sample_data()
    
    print(f"Analysis for: {sample_data['raw_material']}")
    print(f"Quantity needed: {sample_data['quantity_needed']:,.0f} units")
    print(f"Evaluating {len(sample_data['suppliers_data'])} suppliers\n")
    
    # Run the complete analysis
    results = analyze_raw_material_suppliers(
        sample_data['raw_material'],
        sample_data['quantity_needed'], 
        sample_data['suppliers_data']
    )
    
    print("SUPPLIER ANALYSIS RESULTS (Ranked by Reliability Score):")
    print("=" * 60)
    
    for i, result in enumerate(results, 1):
        details = result['details']
        
        print(f"\n{i}. {result['supplier_name']} ({details['country_of_origin']})")
        print(f"   Raw Material: {result['raw_material']}")
        print(f"   Total Cost: ${result['price_after_tariff_for_needed_quantity']:,.2f}")
        print(f"   Reliability Score: {result['reliability_score_after_tariff']:.1f}/100")
        print(f"   ")
        print(f"   Price Details:")
        print(f"     • Base price: ${details['base_price_per_unit']:.2f}/unit")
        print(f"     • After tariffs: ${details['tariffed_price_per_unit']:.2f}/unit")
        print(f"     • Price score: {details['price_score']:.1f}/100")
        print(f"   ")
        print(f"   Quality:")
        print(f"     • Maturity score: {details['maturity_score']:.1f}/100")
        print(f"   ")
        print(f"   Weights used: Price {details['weights_used']['price_weight']:.1f}, "
              f"Maturity {details['weights_used']['maturity_weight']:.1f}")
    
    # Cost comparison
    best_choice = results[0]
    most_expensive = max(results, key=lambda x: x['price_after_tariff_for_needed_quantity'])
    
    print(f"\n" + "=" * 60)
    print("COST COMPARISON:")
    print(f"Best overall choice: {best_choice['supplier_name']} - ${best_choice['price_after_tariff_for_needed_quantity']:,.2f}")
    print(f"Most expensive option: {most_expensive['supplier_name']} - ${most_expensive['price_after_tariff_for_needed_quantity']:,.2f}")
    
    savings = most_expensive['price_after_tariff_for_needed_quantity'] - best_choice['price_after_tariff_for_needed_quantity']
    if savings > 0:
        savings_pct = (savings / most_expensive['price_after_tariff_for_needed_quantity']) * 100
        print(f"Potential savings: ${savings:,.2f} ({savings_pct:.1f}%)")
    
    print(f"\nRecommendation: Choose {best_choice['supplier_name']} for the best balance of cost and quality.")

def parse_arguments():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description='Raw Material Supplier Analysis Tool - Evaluate suppliers based on price, tariffs, and maturity'
    )
    
    parser.add_argument(
        '--input-excel',
        type=str,
        help='Path to input Excel file containing supplier data'
    )
    
    parser.add_argument(
        '--output-excel',
        type=str,
        help='Path for output Excel file with analysis results'
    )
    
    parser.add_argument(
        '--price-weight',
        type=float,
        default=0.6,
        help='Weight for price score in reliability calculation (default: 0.6)'
    )
    
    parser.add_argument(
        '--maturity-weight',
        type=float,
        default=0.4,
        help='Weight for maturity score in reliability calculation (default: 0.4)'
    )
    
    parser.add_argument(
        '--create-sample-excel',
        type=str,
        help='Create a sample Excel input file at the specified path'
    )
    
    parser.add_argument(
        '--demo',
        action='store_true',
        help='Run demonstration with sample data'
    )
    
    return parser.parse_args()

def main():
    """Main entry point of the application."""
    args = parse_arguments()
    
    try:
        # Create sample Excel file if requested
        if args.create_sample_excel:
            logger.info(f"Creating sample Excel file: {args.create_sample_excel}")
            create_sample_excel(args.create_sample_excel)
            return
        
        # Excel-based analysis workflow
        if args.input_excel:
            logger.info(f"Starting Excel-based analysis from: {args.input_excel}")
            
            # Validate weights
            if abs(args.price_weight + args.maturity_weight - 1.0) > 0.001:
                raise ValueError("Price weight and maturity weight must sum to 1.0")
            
            # Determine output file
            output_file = args.output_excel
            if not output_file:
                input_path = Path(args.input_excel)
                output_file = str(input_path.parent / f"{input_path.stem}_analysis.xlsx")
            
            print(f"\n=== EXCEL-BASED SUPPLIER ANALYSIS ===")
            print(f"Input file: {args.input_excel}")
            print(f"Output file: {output_file}")
            print(f"Weights: Price {args.price_weight:.1f}, Maturity {args.maturity_weight:.1f}")
            
            # Run analysis
            results = analyze_from_excel(
                args.input_excel,
                output_file,
                args.price_weight,
                args.maturity_weight
            )
            
            # Display summary
            print(f"\nAnalysis completed successfully!")
            print(f"Evaluated {len(results)} suppliers")
            print(f"Best supplier: {results[0]['supplier_name']}")
            print(f"Best cost: ${results[0]['price_after_tariff_for_needed_quantity']:,.2f}")
            print(f"Reliability score: {results[0]['reliability_score_after_tariff']:.1f}/100")
            print(f"Detailed results saved to: {output_file}")
            return
        
        # Run demo if requested or no arguments provided
        if args.demo or len(sys.argv) == 1:
            main_demo()
            return
        
        # Show help if no valid options provided
        print("Raw Material Supplier Analysis Tool")
        print("\nUsage examples:")
        print("  python main.py --demo                                    # Run demonstration")
        print("  python main.py --create-sample-excel template.xlsx      # Create sample file")
        print("  python main.py --input-excel data.xlsx --output-excel results.xlsx")
        print("\nFor more options, use: python main.py --help")
        
    except Exception as e:
        logger.error(f"Application error: {e}")
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()