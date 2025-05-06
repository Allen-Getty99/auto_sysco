#!/usr/bin/env python3

filename = input("Enter invoice filename: ")
import os
import pandas as pd
import pdfplumber
import re
from collections import defaultdict

# User-changeable paths - modify these variables as needed
PDF_PATH = "FY25 P8 SYSCO 2576717751.pdf"  # Path to the PDF invoice
EXCEL_PATH = "SYSCO_DATABASE.xlsx"  # Path to the Excel database

# Python 3 compatibility note:
# This script is written for Python 3 and requires these packages:
# - pandas: for Excel processing
# - openpyxl: for Excel file reading (used by pandas)
# - pdfplumber: for PDF processing
#
# Setup instructions:
# 1. Navigate to the project directory:
#    cd /Users/allengettyliquigan/Downloads/Project_Auto_GFS
#
# 2. Create a virtual environment:
#    python3 -m venv sysco_env
#
# 3. Activate the virtual environment:
#    - On Mac/Linux: source sysco_env/bin/activate
#    - On Windows: sysco_env\Scripts\activate
#
# 4. Install required packages:
#    pip install pandas openpyxl pdfplumber
#
# 5. Run the script:
#    python sysco_invoice_processor.py

def load_database():
    """Load Excel database and create lookup dictionary."""
    print("Loading database...")
    try:
        # Read the Excel file
        db = pd.read_excel(EXCEL_PATH, sheet_name="Sheet1")
        db.columns = db.columns.str.strip()
        
        # Create a lookup dictionary using item code as key
        db_lookup = {}
        for _, row in db.iterrows():
            # Convert item code to string to handle different formats
            item_code_str = str(row["Item Code"])
            # Extract just the numeric part
            item_code_match = re.search(r'(\d+)', item_code_str)
            if item_code_match:
                item_code = item_code_match.group(1)
                db_lookup[item_code] = {
                    "GL Code": row["GL Code"],
                    "GL Description": row["GL Description"]
                }
        
        print(f"Successfully loaded database with {len(db_lookup)} items.")
        return db_lookup
    except Exception as e:
        print(f"Error loading database: {e}")
        raise

def extract_invoice_data(pdf_path, db_lookup):
    """Extract invoice data using pdfplumber with approach similar to GFS processor."""
    print(f"Processing invoice: {pdf_path}")
    
    invoice_items = []
    bstpz_charges = {"BSTPZ FUEL": 0.0, "BSTPZ DELIVERY SIZE": 0.0, "BSTPZ CREDIT TERMS": 0.0}
    gst_total = 0.0
    
    # Extract all text from PDF
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"
    
    # Process lines
    lines = full_text.split('\n')
    
    for i, line in enumerate(lines):
        # Match line items - looking for 7-digit item code at start of line
        item_match = re.match(r'^\s*(\d{7})\s+', line)
        if item_match:
            item_code_with_zeros = item_match.group(1)
            # Convert to int to drop leading zeros, then back to string
            item_code = str(int(item_code_with_zeros))
            
            # Extract price and total - last two numeric values in the line
            numbers = re.findall(r'\d+\.\d+', line)
            if len(numbers) >= 2:
                price = float(numbers[-2])
                total = float(numbers[-1])
                
                # Match with database - using the version without leading zeros
                if item_code in db_lookup:
                    gl_code = db_lookup[item_code]["GL Code"]
                    gl_desc = db_lookup[item_code]["GL Description"]
                else:
                    gl_code = "Unknown"
                    gl_desc = "Unknown"
                
                # Calculate quantity for reference
                qty = round(total / price, 2) if price != 0 else 0
                
                # Find description (usually in the line itself after the item code)
                # This is approximate since SYSCO invoices vary in format
                desc_match = re.search(r'\d{7}\s+\d+\s+\d+\s+\S+\s+(.+?)\s+\d+\.\d+\s+\d+\.\d+', line)
                description = desc_match.group(1) if desc_match else ""
                
                # Create item data - store the original item code for display
                item = {
                    "Item Code": item_code_with_zeros,  # Keep original with zeros for display
                    "DB Item Code": item_code,          # Version without zeros for reference
                    "Qty": qty,
                    "Price": price,
                    "Total": total,
                    "GL Code": gl_code,
                    "GL Description": gl_desc,
                    "Item Description": description
                }
                
                invoice_items.append(item)
                print(f"Extracted item: {item_code_with_zeros} (DB: {item_code}), Price: {price}, Total: {total}, GL: {gl_desc}")
        
        # Extract BOTTLE DEPOSIT
        if "BOTTLE DEPOSIT" in line and "TOTAL" not in line:
            match = re.search(r"BOTTLE DEPOSIT\s+(\d+\.\d{2})", line)
            if match:
                bd_amount = float(match.group(1))
                bd_item = {
                    "Item Code": "N/A BD",
                    "Qty": 1,
                    "Price": bd_amount,
                    "Total": bd_amount,
                    "GL Code": 600265,
                    "GL Description": "N/A Bev",
                    "Item Description": f"BOTTLE DEPOSIT {bd_amount}"
                }
                invoice_items.append(bd_item)
                print(f"Extracted Bottle Deposit: {bd_amount}")
        
        # Extract RECYCLING FEE
        if "RECYCLING FEE" in line and "TOTAL" not in line:
            match = re.search(r"RECYCLING FEE\s+(\d+\.\d{2})", line)
            if match:
                rf_amount = float(match.group(1))
                rf_item = {
                    "Item Code": "N/A RF",
                    "Qty": 1,
                    "Price": rf_amount,
                    "Total": rf_amount,
                    "GL Code": 600265,
                    "GL Description": "N/A Bev",
                    "Item Description": f"RECYCLING FEE {rf_amount}"
                }
                invoice_items.append(rf_item)
                print(f"Extracted Recycling Fee: {rf_amount}")
        
        # Extract BSTPZ charges
        if "BSTPZ Fuel" in line:
            amt = re.findall(r"\d+\.\d+", line)
            if amt: 
                bstpz_charges["BSTPZ FUEL"] = float(amt[-1])
                print(f"Found BSTPZ FUEL: {bstpz_charges['BSTPZ FUEL']}")
        
        if "BSTPZ Delivery Size" in line:
            amt = re.findall(r"\d+\.\d+", line)
            if amt: 
                bstpz_charges["BSTPZ DELIVERY SIZE"] = float(amt[-1])
                print(f"Found BSTPZ DELIVERY SIZE: {bstpz_charges['BSTPZ DELIVERY SIZE']}")
        
        if "BSTPZ Credit Terms" in line:
            amt = re.findall(r"\d+\.\d+", line)
            if amt: 
                bstpz_charges["BSTPZ CREDIT TERMS"] = float(amt[-1])
                print(f"Found BSTPZ CREDIT TERMS: {bstpz_charges['BSTPZ CREDIT TERMS']}")
        
        # Extract GST information
        if "GST/HST TOTAL" in line or "GST/HST:" in line:
            amt = re.findall(r"\d+\.\d+", line)
            if amt: 
                gst_total = float(amt[-1])
                print(f"Found GST/HST: {gst_total}")
    
    return invoice_items, bstpz_charges, gst_total

def main():
    # Using the predefined PDF path directly without asking
    file_path = PDF_PATH
    
    try:
        # Load database
        db_lookup = load_database()
        
        # Extract data from PDF using pdfplumber
        invoice_items, bstpz_charges, gst_total = extract_invoice_data(file_path, db_lookup)
        
        # Output item table in original extraction order
        print("\n--- Item Table (Original Order) ---")
        print(f"{'Item Code':<12}{'Qty':<10}{'Price':<15}{'Total':<15}{'GL Code':<15}{'GL Description'}")
        print("-" * 85)
        for item in invoice_items:
            # Format price and total with proper spacing
            item_code_str = str(item['Item Code'])
            qty_str = f"{item['Qty']:.2f}"
            price_str = f"{item['Price']:.2f}"
            total_str = f"{item['Total']:.2f}"
            gl_code_str = str(item['GL Code'])
            
            print(f"{item_code_str:<12}{qty_str:<10}{price_str:<15}{total_str:<15}{gl_code_str:<15}{item['GL Description']}")
        
        # Calculate totals by GL Description (for summary only)
        total_by_gl = defaultdict(float)
        
        # First pass - collect all totals
        for item in invoice_items:
            gl_desc = item["GL Description"]
            total_by_gl[gl_desc] += item["Total"]
        
        # Standardize N/A Bev entries (case insensitive)
        standardized_totals = defaultdict(float)
        for gl_desc, total in total_by_gl.items():
            # Normalize all variations of N/A Bev to a single entry
            if gl_desc.upper() in ["N/A BEV", "N/A BEV", "NA BEV"]:
                standardized_totals["N/A Bev"] += total
            else:
                standardized_totals[gl_desc] += total
        
        # GL Description Summary
        print("\n--- GL Description Summary ---")
        for gl_desc, total in standardized_totals.items():
            print(f"{gl_desc}: {round(total, 2)}")
        
        # BSTPZ Summary
        print("\n--- BSTPZ Summary ---")
        bstpz_total = 0
        for key, val in bstpz_charges.items():
            print(f"{key}: {round(val, 2)}")
            bstpz_total += val
        print(f"BSTPZ total: {round(bstpz_total, 2)}")
        
        # GST
        print(f"\nGST/HST: {round(gst_total, 2)}")
        
        # Grand Total
        grand_total = sum(standardized_totals.values()) + bstpz_total + gst_total
        print(f"\nGrand Total: ${grand_total:.2f}")
        
    except Exception as e:
        print(f"Error processing invoice: {e}")
        import traceback
        traceback.print_exc()  # Print detailed error information

if __name__ == "__main__":
    main()
