import pandas as pd
import os

class InputHandler:
    """
    Handles reading and validating the Excel input file.
    """

    REQUIRED_COLUMNS = ['File_Name', 'Chapter_Name', 'Topic_Name', 'Class']

    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.dataframe = None

    def read_excel(self):
        """
        Reads Excel file with only required columns.
        """
        if not os.path.isfile(self.excel_path):
            raise FileNotFoundError(f"[InputHandler] ERROR: Excel file not found at {self.excel_path}")

        try:
            self.dataframe = pd.read_excel(self.excel_path, usecols=self.REQUIRED_COLUMNS)
            print(f"[InputHandler] Loaded Excel with {len(self.dataframe)} rows.")
        except Exception as e:
            raise ValueError(f"[InputHandler] ERROR reading Excel: {e}")

    def validate_columns(self):
        """
        Ensures required columns are present and no missing values.
        """
        missing_cols = [col for col in self.REQUIRED_COLUMNS if col not in self.dataframe.columns]
        if missing_cols:
            raise ValueError(f"[InputHandler] ERROR: Missing columns → {missing_cols}")

        
    def extract_file_list(self):
        """
        Converts dataframe to list of dictionaries (records) with normalized keys.
        """
        # Rename columns to lowercase keys
        normalized_df = self.dataframe.rename(columns={
            'File_Name': 'file_name',
            'Chapter_Name': 'chapter_name',
            'Topic_Name': 'topic_name',
            'Class': 'class_name'
        })

        # Convert to list of dicts
        file_entries = normalized_df.to_dict(orient='records')
        print(f"[InputHandler] Extracted {len(file_entries)} file entries.")
        return file_entries

