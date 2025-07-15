import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

class EmployeeCostMapper:
    def __init__(self, target_path, processed_path, target_sheet='Project', skip_names=None):
        self.target_path = target_path
        self.processed_path = processed_path
        self.target_sheet = target_sheet
        self.processed_df = pd.read_excel(self.processed_path)
        self.wb = load_workbook(self.target_path)
        self.ws = self.wb[self.target_sheet]
        self.month_col_map = self._map_month_columns()
        self.operation_names = set(self.processed_df['Operation'].unique())
        self.processed_df['Name_norm'] = self.processed_df['Name'].apply(self.normalize_name)
        self.processed_df['Operation_norm'] = self.processed_df['Operation'].str.strip()
        self.skip_names = set(n.strip().lower() for n in (skip_names or []))
        self.mapping = []

    def _map_month_columns(self):
        """Map month names to their column indices."""
        month_col_map = {}
        for col in range(6, self.ws.max_column + 1):  # Month columns start at F (no.6)
            cell_value = self.ws.cell(row=1, column=col).value
            if cell_value:
                # Normalize to 'YYYY-MM' string
                if isinstance(cell_value, str) and cell_value[:4].isdigit():
                    period = cell_value[:7]
                elif hasattr(cell_value, 'strftime'):
                    period = cell_value.strftime('%Y-%m')
                else:
                    continue
                month_col_map[period] = col
        return month_col_map

    def _is_operation(self, value):
        return value and str(value).strip() in self.operation_names

    def _should_skip(self, value):
        return value and str(value).strip().lower() in self.skip_names

    def map_employees(self):
        current_operation = None
        for row in range(2, self.ws.max_row + 1):
            cell_value = self.ws.cell(row=row, column=5).value  # Column E

            # Stop processing if "Accumulated Total" is reached
            if cell_value and str(cell_value).strip().lower() == "accumulated total":
                print(f"Reached 'Accumulated Total' at row {row}. Stopping iteration.")
                break

            # Detect operation section
            if self._is_operation(cell_value):
                current_operation = str(cell_value).strip()
                continue

            # Skip if not in an operation section
            if not current_operation:
                continue

            emp_name = self.normalize_name(cell_value)
            if not emp_name or emp_name.strip() == '' or self._is_operation(emp_name) or self._should_skip(emp_name):
                continue

            # For each month, get current value and cell address
            for period, col in self.month_col_map.items():
                cell = self.ws.cell(row=row, column=col)
                cell_address = f"{get_column_letter(col)}{row}"
                self.mapping.append({
                    'Operation': current_operation,
                    'Employee': emp_name.strip(),
                    'Month': period,
                    'Cell': cell_address,
                    'CurrentValue': cell.value
                })

    def normalize_name(self,name):
        if not name:
            return None
        return ALIAS_TO_CANONICAL.get(name.strip().lower(), name.strip())


    def update_costs(self):
        orange_font = Font(color="FFA500")
        update_count = 0
        for entry in self.mapping:
            # Find the correct fee in the processed DataFrame
            mask = (
                (self.processed_df['Operation'] == entry['Operation']) &
                (self.processed_df['Name_norm'] == entry['Employee']) &
                (self.processed_df['StartDate'].str.startswith(entry['Month']))
            )
            match = self.processed_df[mask]
            if not match.empty:
                fee = match['Fees'].values[0]
                cell = self.ws[entry['Cell']]
                cell.value = fee
                cell.font = orange_font
                update_count += 1
        print(f"Updated {update_count} cells with new costs (highlighted in orange).")
        # Save the workbook
        self.wb.save(self.target_path.replace('.xlsx', '_updated.xlsx'))


if __name__ == "__main__":
    TARGET_FILE = "../docs/target_sheet.xlsx"
    PROCESSED_FILE = "../docs/processed_timesheet.xlsx"
    SHEET_NAME = "Project"

    # rows to skip 
    SKIP_NAMES = [
        "Test - ARE 5240", "Service Management - ARE 5290 + int.", "JCC",
        "CHCM (Eviden)", "AWS Encryption Key (KMS for Eightfold; GBS)",
        "DirX (Ext)/(SAG Global)", "Integrations DPS", "GBS total",
        "TRE (Tupu) PO", "Eightfold (PO Q1/2025: 9708791569; Q2-Q4: tbd.)",
        "Provider", "CERT check (pen GBS)", "Accessibility Test",
        "Travel & Hospitality", "Additional costs", "Total Costs",
        "Accumulated costs Eightfold Crew", "Service Management - ARE 5290 + int.",
        "DirX PT -  5240", "DirX (Eviden)", "TRE (Tupu) (PO tbd)", "Eightfold",
        "Avature DT (PO 9708872111)", "Avature AM (PO 9708872111)",
        "Avature Healthcheck (PO", "Total costs", "Accumulated costs Avature Crew",
        "CHCM (Evdien)", "Avature ext. Careers Portal (PO 9709043748) - PDP",
        "Avature ext. Careers Portal (PO 9709132497)", "Travel & Hospitality ext. Provider",
        "Travel & Hospitality DE", "Travel & Hospitality ES / CZ / PT",
        "Total costs", "Accumulated costs Ext. Careers Port Crew", "Accumulated costs Preboarding",
        "Service Management","Total","Avature","DPS Internal Umlage (global)"
    ]

    NAME_MAP = {
    "Assuncao Gambetta Clemente, Fernanda": [
        "Assuncao Gambetta Clemente, Fernanda",
        "Fernanda Assuncao Gambetta Clemente",
    ],
    "Magro, Daniel": [
        "Magro, Daniel",
    ],
    "Pires Rosa, Claudia": [
        "Pires Rosa, Claudia",
        "Claudia Pires Rosa",
        "Rosa, Claudia (ext)",
        "Pires Rosa, Claudia (ext)",
        "Claudia Rosa",
    ],
    "Helbing, Björn": [
        "Helbing, Björn",
        "Helbing, Bjoern",
        "B. Helbing",
        "Björn Helbing",
    ],
    "Matos Oliveira, Ana Rita": [
        "Matos Oliveira, Ana Rita",
        "Ana Rita Matos Oliveira",
        "Matos dos Santos Oliveira, Ana Rita",
        "Matos Oliveira, Rita",
    ],
    "Pires, Filipe": [
        "Pires, Filipe",
        "Guerreiro Luis Pires, Filipe Viegas",
        "Filipe Pires",
    ],
    "Plácido, Andreia": [
        "Plácido, Andreia",
        "Moreira Cristo Placido, Andreia Sofia",
        "Andreia Plácido",
        "Andreia Sofia Moreira Cristo Placido",
    ],
    "Antunes, Ricardo": [
        "Antunes, Ricardo",
        "Ricardo Antunes",
    ],
    "Fernandes Redondo, Amanda": [
        "Fernandes Redondo, Amanda",
        "Amanda Fernandes Redondo",
        "Fernandez Redondo, Amanda",
    ],
    "Zouine, Meryem": [
        "Zouine, Meryem",
        "Meryem Zouine",
    ],
    "Cerezo, Alberto": [
        "Cerezo, Alberto",
        "Cerezo Ruiz, Alberto",
        "Alberto Cerezo",
    ],
    "Swoboda, Claudia": [
        "Swoboda, Claudia",
        "Claudia Swoboda",
    ],
    "Bicho, Rita": [
        "Bicho, Rita",
        "Lazaro Bicho, Rita Sofia",
        "Rita Bicho",
        "Rita Sofia Lazaro Bicho",
    ],
    "Lopes Fonseca, Mario Andre": [
        "Lopes Fonseca, Mario Andre",
        "Mario Andre Lopes Fonseca",
    ],
    "Wiesheu, Andreas": [
        "Wiesheu, Andreas",
        "Andreas Wiesheu",
    ],
    "Heldwein, Christian": [
        "Heldwein, Christian",
        "Christian Heldwein",
    ],
    "Hernandes Vaz, Joao Rafael": [
        "Hernandes Vaz, Joao Rafael",
        "Joao Rafael Hernandes Vaz",
    ],
    "Vitorino, Diana": [
        "Vitorino, Diana",
        "Diana Vitorino",
    ],
    "Candeias Gracioso, Sara Margarida": [
        "Candeias Gracioso, Sara Margarida",
        "Sara Margarida Candeias Gracioso",
    ],
    "do Nascimento Matos Manso, Rui Pedro": [
        "do Nascimento Matos Manso, Rui Pedro",
        "Rui Pedro do Nascimento Matos Manso",
    ],
}
ALIAS_TO_CANONICAL = {}
for canonical, aliases in NAME_MAP.items():
    for alias in aliases:
        ALIAS_TO_CANONICAL[alias.strip().lower()] = canonical
    
mapper = EmployeeCostMapper(TARGET_FILE, PROCESSED_FILE, SHEET_NAME, SKIP_NAMES)
mapper.map_employees()
mapper.update_costs()