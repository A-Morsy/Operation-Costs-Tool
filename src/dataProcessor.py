import locale
import pandas as pd
import math
from openpyxl import load_workbook

locale.setlocale(locale.LC_ALL, 'de_DE.UTF-8')

class TimeSheetUpdater:
    def __init__(self):
        self.source_data = None
        self.target_data = None
        
        wb = load_workbook('../docs/target_sheet.xlsx')
        # Updated work package mappings
        self.work_package_to_operation = {
            # Avature Crew packages
            '001 / Avature Crew  - PM/CM': 'Avature Crew',
            '003 / Avature Crew  - Testcenter': 'Avature Crew',
            '004 / Avature Crew  - Integrations': 'Avature Crew',
            
            # Avature Preboarding packages
            '052 / Avature Crew - Pre-/Onboarding': 'Avature Preboarding',
            
            # Eightfold Crew packages
            '007 / Eightfold Crew  - Testcenter': 'Eightfold Crew',
            '005 / Eightfold Crew - PM/CM': 'Eightfold Crew',
            '008 / Eightfold Crew  - Integrations': 'Eightfold Crew',
            
            # Avature ext. Careers Portal Crew packages
            '053 / Ext.Careers Portal Crew - PM/CM': 'Avature ext. Careers Portal Crew',
            '054 / Ext.Careers Portal Crew - Integration': 'Avature ext. Careers Portal Crew',
            '056 / Ext.Careers Portal Crew - Testcenter': 'Avature ext. Careers Portal Crew'
        }

    def custom_round(self,value):
        decimal_part = value - int(value)
        result = math.ceil(value) if decimal_part >= 0.5 else math.floor(value)
        return result
    
    def format_period(self, date):
        """Convert date to 'MMM-YY' format (e.g., 'Oct-24')"""
        try:
            if isinstance(date, str):
                # Try to parse the date string
                date = pd.to_datetime(date)
            return date.strftime('%b-%y').title()
        except:
            return date

    def load_and_process_data(self, source_path):
        """Load and process source data into desired format"""
        # Read source data
        self.source_data = pd.read_excel(source_path)
        
        # Add operation column based on work package mapping
        self.source_data['Operation'] = self.source_data['Work Package'].map(self.work_package_to_operation)
        
        null_count = self.source_data['Operation'].isna().sum()
        print(f"Number of null operation rows:{null_count}")
        
        # Drop rows where Operation is None (unmapped work packages)
        self.source_data = self.source_data.dropna(subset=['Operation'])
        
        # Convert Period to datetime if it isn't already
        self.source_data['Period'] = pd.to_datetime(self.source_data['Period'])
        
        # Format Period to 'MMM-YY'
        self.source_data['Period'] = self.source_data['Period'].apply(self.format_period)
        
        # Group by operation, name, and period
        processed_data = self.source_data.groupby(
            ['Operation', 'Name', 'Period']
        ).agg({
            'Fees': 'sum'
        }).reset_index()
        
        processed_data['StartDate'] = (pd.to_datetime(
            processed_data['Period'].apply(
                lambda x: f"01-{x}"
            ), 
            format='%d-%b-%y'
        ).apply(lambda x: x.strftime('%Y-%m')))

        processed_data['Fees'] = processed_data['Fees'].apply(lambda x: self.custom_round(x))
        

        # Sort the data
        processed_data = processed_data.sort_values(
            ['Operation', 'Name', 'StartDate'],
            ascending=[False,True,True]
        ).reset_index(drop=True)
        
        return processed_data

    def save_processed_data(self, data, output_path):
        """Save the processed data to Excel"""
        # Save to Excel with specific formatting
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            data.to_excel(
                writer,
                index=False,
                sheet_name='Processed Data'
            )
            
            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Processed Data']
            
            # Adjust column widths
            worksheet.column_dimensions['A'].width = 40  # Operation
            worksheet.column_dimensions['B'].width = 35  # Name
            worksheet.column_dimensions['C'].width = 12  # Period
            worksheet.column_dimensions['D'].width = 12  # Fees
            
        print(f"\nProcessed data saved to: {output_path}")
        print("\nFirst few rows of processed data:")
        print(data.head())
        
        print("\nUnique periods in the data:")
        print(sorted(data['Period'].unique()))

def main():
    updater = TimeSheetUpdater()
    
    try:
        # Process the source data
        print("Processing data...")
        processed_data = updater.load_and_process_data('../docs/source_timesheet.xlsx')
        
        # Save the processed data
        print("\nSaving processed data...")
        updater.save_processed_data(processed_data, '../docs/processed_timesheet.xlsx')
        
        print("\nProcess completed successfully!")
        
    except Exception as e:
        print(f"\nError: {str(e)}")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    main()