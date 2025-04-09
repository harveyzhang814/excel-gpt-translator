import unittest
from pathlib import Path
import pandas as pd
from src.core.translator import Translator
from src.core.config import Config

class TestTranslator(unittest.TestCase):
    def setUp(self):
        self.config = Config()
        self.translator = Translator(self.config)
        
        # Create a test Excel file
        self.test_file = Path("test.xlsx")
        df = pd.DataFrame({
            "Column1": ["Hello", "World", "Test", "Data"],
            "Column2": ["Test", "Data", "Hello", "World"]
        })
        df.to_excel(self.test_file, index=False)
    
    def tearDown(self):
        # Clean up test files
        if self.test_file.exists():
            self.test_file.unlink()
        for file in Path(".").glob("test_*.xlsx"):
            file.unlink()
    
    def test_translate_excel(self):
        task_data = {
            "file": str(self.test_file),
            "sheet": "Sheet1",
            "cell_range": "A1:B2",
            "current_language": "English",
            "target_languages": ["Spanish"],
            "comparison_mode": False,
            "prompt": "You are a professional translator. Translate the following text from {current_lang} to {target_lang}. Maintain the original meaning and tone."
        }
        
        self.translator.translate_excel(task_data)
        
        # Check if output file was created
        output_file = Path("test_Spanish.xlsx")
        self.assertTrue(output_file.exists())
        
        # Check if the file can be read and has the correct number of rows
        df = pd.read_excel(output_file)
        self.assertEqual(len(df), 2)  # Should only have 2 rows due to cell range
        self.assertEqual(len(df.columns), 2)
    
    def test_invalid_cell_range(self):
        task_data = {
            "file": str(self.test_file),
            "sheet": "Sheet1",
            "cell_range": "InvalidRange",
            "current_language": "English",
            "target_languages": ["Spanish"],
            "comparison_mode": False,
            "prompt": "You are a professional translator. Translate the following text from {current_lang} to {target_lang}. Maintain the original meaning and tone."
        }
        
        # Should not raise an exception
        self.translator.translate_excel(task_data)
        
        # Check if output file was created with all rows
        output_file = Path("test_Spanish.xlsx")
        self.assertTrue(output_file.exists())
        
        df = pd.read_excel(output_file)
        self.assertEqual(len(df), 4)  # Should have all rows due to invalid range 