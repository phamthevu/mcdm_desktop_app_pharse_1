from PyQt6.QtWidgets import QMainWindow, QTabWidget
from ui.tab_config import TabConfig

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MCDM Decision Support System (Fuzzy & Z-Number)")
        self.resize(1100, 720)

        # Tab Widget container
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Tab 1: Configuration
        self.tab_config = TabConfig()
        self.tabs.addTab(self.tab_config, "Configuration (Setup)")

        # Hook START event
        self.tab_config.started.connect(self.on_survey_started)

    def on_survey_started(self, config):
        print(f"[START] Survey initialized for model: {config.model_type} (Z-Number: {config.use_z_number})")