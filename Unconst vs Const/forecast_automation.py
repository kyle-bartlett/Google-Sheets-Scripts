#!/usr/bin/env python3
"""
ANKER FORECAST AUTOMATION - PYTHON VERSION
This script transforms wide forecast data to tall format and uploads to Google Sheets
Alternative to Apps Script for more reliable automation
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os
import sys

# Google Sheets integration (optional)
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    GOOGLE_SHEETS_AVAILABLE = False
    print("Google Sheets integration not available. Install gspread and google-auth for full automation.")

class ForecastAutomation:
    def __init__(self, excel_file_path):
        self.excel_file = excel_file_path
        self.data = {}
        self.output_data = None
        
    def load_data(self):
        """Load data from Excel file"""
        print(f"Loading data from {self.excel_file}")
        
        # Load the wide format sheets
        try:
            self.data['constrained'] = pd.read_excel(self.excel_file, sheet_name='Constrained Wide')
            self.data['unconstrained'] = pd.read_excel(self.excel_file, sheet_name='Unconstrained Wide')
            print(f"✓ Loaded Constrained Wide: {self.data['constrained'].shape}")
            print(f"✓ Loaded Unconstrained Wide: {self.data['unconstrained'].shape}")
        except Exception as e:
            print(f"Error loading data: {e}")
            return False
            
        return True
    
    def get_quarter(self, week):
        """Convert week number to quarter"""
        week = int(week)
        if 202529 <= week <= 202539:
            return 'Q3 2025'
        elif 202540 <= week <= 202552:
            return 'Q4 2025'
        elif 202553 <= week <= 202605:
            return 'Q1 2026'
        else:
            return 'Unknown'
    
    def transform_to_tall(self):
        """Transform wide format data to tall format"""
        print("Transforming data from wide to tall format...")
        
        # Identify week columns (format: 202xxx)
        constrained_df = self.data['constrained']
        week_columns = []
        
        for col in constrained_df.columns:
            col_str = str(col)
            if col_str.isdigit() and col_str.startswith('202') and len(col_str) == 6:
                week_columns.append(col)  # Keep original column type (int or str)
        
        print(f"Found {len(week_columns)} week columns: {week_columns[:5]}..." if len(week_columns) > 5 else week_columns)
        
        if not week_columns:
            raise ValueError("No week columns found. Expected columns with format 202xxx")
        
        # Create unconstrained lookup for fast access
        unconstrained_lookup = {}
        unconstrained_df = self.data['unconstrained']
        
        for _, row in unconstrained_df.iterrows():
            helper = str(row.iloc[0])  # Helper column
            # Handle different column names between sheets
            sell_in_price_col = 'Sell-in Price' if 'Sell-in Price' in unconstrained_df.columns else 'Sell-in price'
            sell_in_price = float(row[sell_in_price_col]) if pd.notna(row[sell_in_price_col]) else 0
            
            for week in week_columns:
                if week in unconstrained_df.columns:
                    units = float(row[week]) if pd.notna(row[week]) else 0
                    revenue = units * sell_in_price
                    
                    key = f"{helper}_{str(week)}"
                    unconstrained_lookup[key] = {
                        'units': units,
                        'revenue': revenue
                    }
        
        print(f"Created unconstrained lookup with {len(unconstrained_lookup)} entries")
        
        # Transform constrained data
        output_rows = []
        
        processed_rows = 0
        for _, row in constrained_df.iterrows():
            helper = str(row.iloc[0])
            # Handle different column names between sheets
            sell_in_price_col = 'Sell-in Price' if 'Sell-in Price' in constrained_df.columns else 'Sell-in price'
            sell_in_price = float(row[sell_in_price_col]) if pd.notna(row[sell_in_price_col]) else 0
            pct = row.iloc[2]
            pdt = row.iloc[3]
            customer_id = row.iloc[4]
            customer = row.iloc[5]
            sku = row.iloc[6]
            sku_description = row.iloc[7]
            
            # Skip rows with missing essential data
            if pd.isna(helper) or pd.isna(customer) or pd.isna(sku) or helper == 'nan':
                continue
                
            processed_rows += 1
            if processed_rows <= 3:  # Debug first few rows
                print(f"Processing row {processed_rows}: {helper}, {customer}, {sku}")
            
            for week in week_columns:
                if week in constrained_df.columns:
                    constrained_units = float(row[week]) if pd.notna(row[week]) else 0
                    constrained_revenue = constrained_units * sell_in_price
                    
                    # Look up unconstrained data
                    key = f"{helper}_{str(week)}"
                    unconstrained = unconstrained_lookup.get(key, {'units': 0, 'revenue': 0})
                    
                    delta_units = constrained_units - unconstrained['units']
                    delta_revenue = constrained_revenue - unconstrained['revenue']
                    
                    quarter = self.get_quarter(int(week))
                    is_current_q = quarter == 'Q4 2025'
                    gap_flag = 'Supply Gap' if delta_units < 0 else ''
                    
                    # Add constrained row
                    output_rows.append({
                        'Customer': customer,
                        'Anker SKU': sku,
                        'PDT': pdt,
                        'Forecast Type': 'Constrained',
                        'Quarter': quarter,
                        'Week': int(week),
                        'Forecast - Units': constrained_units,
                        'Forecast Revenue': constrained_revenue,
                        'Delta Units': delta_units,
                        'Delta - Revenue': delta_revenue,
                        'Gap Flag': gap_flag,
                        'IsCurrentQ': is_current_q,
                        'Helper': helper,
                        'Sell-In Price': sell_in_price
                    })
                    
                    # Add unconstrained row
                    output_rows.append({
                        'Customer': customer,
                        'Anker SKU': sku,
                        'PDT': pdt,
                        'Forecast Type': 'Unconstrained',
                        'Quarter': quarter,
                        'Week': int(week),
                        'Forecast - Units': unconstrained['units'],
                        'Forecast Revenue': unconstrained['revenue'],
                        'Delta Units': 0,  # Delta is 0 for unconstrained (baseline)
                        'Delta - Revenue': 0,
                        'Gap Flag': '',
                        'IsCurrentQ': is_current_q,
                        'Helper': helper,
                        'Sell-In Price': sell_in_price
                    })
        
        self.output_data = pd.DataFrame(output_rows)
        print(f"✓ Processed {processed_rows} data rows")
        print(f"✓ Transformation complete! Created {len(self.output_data)} records")
        
        return self.output_data
    
    def create_summaries(self):
        """Create summary DataFrames for dashboards"""
        if self.output_data is None:
            raise ValueError("No output data available. Run transform_to_tall() first.")
        
        summaries = {}
        
        # Filter only supply gaps
        gaps_df = self.output_data[self.output_data['Gap Flag'] == 'Supply Gap'].copy()
        
        # SKU Summary
        sku_summary = gaps_df.groupby(['Anker SKU', 'PDT']).agg({
            'Delta Units': 'sum',
            'Delta - Revenue': 'sum',
            'Customer': 'nunique'
        }).reset_index()
        sku_summary.columns = ['SKU', 'PDT', 'Gap Units', 'Revenue Impact', 'Customers Affected']
        sku_summary['Gap Units'] = sku_summary['Gap Units'].abs()
        sku_summary['Revenue Impact'] = sku_summary['Revenue Impact'].abs()
        sku_summary = sku_summary.sort_values('Revenue Impact', ascending=False)
        sku_summary.insert(0, 'Rank', range(1, len(sku_summary) + 1))
        summaries['sku_summary'] = sku_summary
        
        # Customer Summary
        customer_summary = gaps_df.groupby('Customer').agg({
            'Delta Units': 'sum',
            'Delta - Revenue': 'sum',
            'Anker SKU': 'nunique'
        }).reset_index()
        customer_summary.columns = ['Customer', 'Gap Units', 'Revenue Impact', 'SKUs Affected']
        customer_summary['Gap Units'] = customer_summary['Gap Units'].abs()
        customer_summary['Revenue Impact'] = customer_summary['Revenue Impact'].abs()
        customer_summary = customer_summary.sort_values('Revenue Impact', ascending=False)
        customer_summary.insert(0, 'Rank', range(1, len(customer_summary) + 1))
        summaries['customer_summary'] = customer_summary
        
        # Weekly Trends
        weekly_trends = gaps_df.groupby(['Week', 'Quarter']).agg({
            'Delta Units': 'sum',
            'Delta - Revenue': 'sum',
            'Customer': 'count'
        }).reset_index()
        weekly_trends.columns = ['Week', 'Quarter', 'Gap Units', 'Revenue Impact', 'Records Count']
        weekly_trends['Gap Units'] = weekly_trends['Gap Units'].abs()
        weekly_trends['Revenue Impact'] = weekly_trends['Revenue Impact'].abs()
        weekly_trends = weekly_trends.sort_values('Week')
        summaries['weekly_trends'] = weekly_trends
        
        # PDT Summary
        pdt_summary = gaps_df.groupby('PDT').agg({
            'Delta Units': 'sum',
            'Delta - Revenue': 'sum',
            'Customer': 'nunique'
        }).reset_index()
        pdt_summary.columns = ['PDT', 'Gap Units', 'Revenue Impact', 'Customers Affected']
        pdt_summary['Gap Units'] = pdt_summary['Gap Units'].abs()
        pdt_summary['Revenue Impact'] = pdt_summary['Revenue Impact'].abs()
        pdt_summary = pdt_summary.sort_values('Revenue Impact', ascending=False)
        summaries['pdt_summary'] = pdt_summary
        
        print("✓ Created summary tables:")
        for name, df in summaries.items():
            print(f"  - {name}: {len(df)} rows")
        
        return summaries
    
    def save_to_excel(self, output_file='forecast_analysis_output.xlsx'):
        """Save all data to Excel file"""
        if self.output_data is None:
            raise ValueError("No output data available. Run transform_to_tall() first.")
        
        summaries = self.create_summaries()
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Main data
            self.output_data.to_excel(writer, sheet_name='Looker_Ready_View', index=False)
            
            # Summaries
            summaries['sku_summary'].to_excel(writer, sheet_name='SKU_Summary', index=False)
            summaries['customer_summary'].to_excel(writer, sheet_name='Customer_Summary', index=False)
            summaries['weekly_trends'].to_excel(writer, sheet_name='Weekly_Trends', index=False)
            summaries['pdt_summary'].to_excel(writer, sheet_name='PDT_Summary', index=False)
        
        print(f"✓ Saved analysis to {output_file}")
        return output_file
    
    def upload_to_google_sheets(self, spreadsheet_id, credentials_file=None):
        """Upload data to Google Sheets (requires service account credentials)"""
        if not GOOGLE_SHEETS_AVAILABLE:
            print("❌ Google Sheets integration not available. Install gspread and google-auth")
            return False
        
        if self.output_data is None:
            raise ValueError("No output data available. Run transform_to_tall() first.")
        
        # Set up credentials
        scopes = ['https://www.googleapis.com/auth/spreadsheets']
        
        if credentials_file and os.path.exists(credentials_file):
            creds = Credentials.from_service_account_file(credentials_file, scopes=scopes)
        else:
            print("❌ Credentials file not found. Cannot upload to Google Sheets")
            return False
        
        try:
            client = gspread.authorize(creds)
            spreadsheet = client.open_by_key(spreadsheet_id)
            
            # Upload main data
            try:
                worksheet = spreadsheet.worksheet('Looker_Ready_View_Python')
            except:
                worksheet = spreadsheet.add_worksheet('Looker_Ready_View_Python', 
                                                    rows=len(self.output_data) + 100, 
                                                    cols=len(self.output_data.columns))
            
            worksheet.clear()
            worksheet.update([self.output_data.columns.values.tolist()] + self.output_data.values.tolist())
            
            print("✓ Uploaded to Google Sheets successfully")
            return True
            
        except Exception as e:
            print(f"❌ Error uploading to Google Sheets: {e}")
            return False
    
    def print_summary_stats(self):
        """Print key statistics"""
        if self.output_data is None or len(self.output_data) == 0:
            print("No data available")
            return
        
        gaps_df = self.output_data[self.output_data['Gap Flag'] == 'Supply Gap']
        
        print("\n" + "="*50)
        print("FORECAST ANALYSIS SUMMARY")
        print("="*50)
        print(f"Total records processed: {len(self.output_data):,}")
        print(f"Supply gap records: {len(gaps_df):,}")
        print(f"Total revenue at risk: ${gaps_df['Delta - Revenue'].abs().sum():,.2f}")
        print(f"Total units at risk: {gaps_df['Delta Units'].abs().sum():,.0f}")
        print(f"Unique SKUs affected: {gaps_df['Anker SKU'].nunique()}")
        print(f"Unique customers affected: {gaps_df['Customer'].nunique()}")
        
        # Quarter breakdown
        print("\nQuarter breakdown:")
        quarter_summary = gaps_df.groupby('Quarter')['Delta - Revenue'].sum().abs()
        for quarter, revenue in quarter_summary.items():
            print(f"  {quarter}: ${revenue:,.2f}")


def main():
    """Main execution function"""
    excel_file = 'Unconstrained FCST vs Constrained FCST 25WK29 FC Version.xlsx'
    
    if not os.path.exists(excel_file):
        print(f"❌ Excel file not found: {excel_file}")
        print("Make sure the file is in the current directory")
        return
    
    print("🚀 Starting Anker Forecast Automation")
    print(f"📁 Processing file: {excel_file}")
    
    # Initialize automation
    automation = ForecastAutomation(excel_file)
    
    try:
        # Load and transform data
        if not automation.load_data():
            return
        
        automation.transform_to_tall()
        automation.print_summary_stats()
        
        # Save to Excel
        output_file = automation.save_to_excel()
        
        print(f"\n✅ Automation complete!")
        print(f"📊 Output saved to: {output_file}")
        print("\nNext steps:")
        print("1. Upload the output Excel file to Google Drive")
        print("2. Open in Google Sheets")
        print("3. Create pivot tables and charts from the Looker_Ready_View sheet")
        print("4. Or use the fixed Apps Script for direct Google Sheets automation")
        
    except Exception as e:
        print(f"❌ Error during automation: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
