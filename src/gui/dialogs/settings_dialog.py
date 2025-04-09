from PyQt6.QtWidgets import QDialog, QFormLayout, QLineEdit, QHBoxLayout, QPushButton

class SettingsDialog(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.setWindowTitle("Settings")
        self.setup_ui()
    
    def setup_ui(self):
        layout = QFormLayout(self)
        
        # API Key
        self.api_key = QLineEdit()
        self.api_key.setText(self.config.get_api_key())
        layout.addRow("OpenAI API Key:", self.api_key)
        
        # Buttons
        buttons = QHBoxLayout()
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        buttons.addWidget(save_btn)
        buttons.addWidget(cancel_btn)
        layout.addRow("", buttons)
        
        save_btn.clicked.connect(self.save_settings)
        cancel_btn.clicked.connect(self.reject)
    
    def save_settings(self):
        self.config.save_api_key(self.api_key.text())
        self.accept() 