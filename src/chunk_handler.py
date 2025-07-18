# chunk_handler.py
import pandas as pd

class ChunkHandler:
    def __init__(self, df, normalize_name_func, operation_col='Operation', name_col='Name', date_col='StartDate'):
        """
        :param df: pandas DataFrame to chunk
        :param normalize_name_func: function to normalize names
        :param operation_col: column name for operation/section
        :param name_col: column name for employee names
        :param date_col: column name for dates
        """
        self.df = df.copy()
        self.operation_col = operation_col
        self.name_col = name_col
        self.date_col = date_col
        self.normalize_name = normalize_name_func
        
        if self.name_col in self.df.columns:
            self.df['Name_norm'] = self.df[self.name_col].apply(self.normalize_name)
        else:
            self.df['Name_norm'] = None
    
    def add_normalized_columns(self):
        # Normalize operation, strip whitespace
        if self.operation_col in self.df.columns:
            self.df['Operation_norm'] = self.df[self.operation_col].astype(str).str.strip()
        else:
            self.df['Operation_norm'] = None
        
        if self.date_col in self.df.columns:
            self.df[self.date_col] = pd.to_datetime(self.df[self.date_col], errors='coerce')
            self.df['Month'] = self.df[self.date_col].dt.strftime('%Y-%m')
        else:
            self.df['Month'] = None
    
    def get_chunks_by_operation(self):
        """Yield tuples (operation, chunk_df) for each unique operation"""
        self.add_normalized_columns()
        unique_ops = self.df['Operation_norm'].dropna().unique()
        for op in unique_ops:
            chunk_df = self.df[self.df['Operation_norm'] == op].copy()
            yield op, chunk_df

    def save_chunks(self, out_dir, filename_prefix='chunk'):
        import os
        os.makedirs(out_dir, exist_ok=True)
        for op, chunk in self.get_chunks_by_operation():
            # Remove rows with any NaN values before saving
            chunk_clean = chunk.dropna()
            safe_op = op.replace(' ', '_').replace('/', '_')
            filename = f"{filename_prefix}_{safe_op}.xlsx"
            path = os.path.join(out_dir, filename)
            chunk_clean.to_excel(path, index=False)
            print(f"Saved chunk for operation '{op}' with {len(chunk_clean)} rows to {path}")
