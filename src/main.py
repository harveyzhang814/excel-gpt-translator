import sys
from PyQt6.QtWidgets import QApplication
from gui.main_window import MainWindow
from core.config import Config

def main():
    app = QApplication(sys.argv)
    
    # Initialize configuration
    config = Config()
    
    # Create and show the main window
    window = MainWindow(config)
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 