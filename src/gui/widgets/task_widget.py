from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QProgressBar

class TaskWidget(QWidget):
    def __init__(self, task_id, task_data, parent=None):
        super().__init__(parent)
        self.task_id = task_id
        self.task_data = task_data
        self.setup_ui()
    
    def setup_ui(self):
        layout = QHBoxLayout(self)
        
        # Task info
        info_layout = QVBoxLayout()
        self.file_label = QLabel(f"File: {self.task_data['file']}")
        self.sheet_label = QLabel(f"Sheet: {self.task_data['sheet']}")
        self.target_label = QLabel(f"Target: {', '.join(self.task_data['target_languages'])}")
        info_layout.addWidget(self.file_label)
        info_layout.addWidget(self.sheet_label)
        info_layout.addWidget(self.target_label)
        layout.addLayout(info_layout)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        # Buttons layout
        buttons_layout = QVBoxLayout()
        
        # Start button
        self.start_btn = QPushButton("Start")
        buttons_layout.addWidget(self.start_btn)
        
        # Edit button
        self.edit_btn = QPushButton("Edit")
        buttons_layout.addWidget(self.edit_btn)
        
        # Remove button
        self.remove_btn = QPushButton("Remove")
        buttons_layout.addWidget(self.remove_btn)
        
        layout.addLayout(buttons_layout) 