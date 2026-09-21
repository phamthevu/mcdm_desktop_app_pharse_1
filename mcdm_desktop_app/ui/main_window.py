from PyQt6.QtWidgets import QMainWindow, QTabWidget
from ui.tab_config import TabConfig
from ui.tab_survey import TabSurvey

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MCDM Decision Support System (Fuzzy & Z-Number)")
        self.resize(1150, 750)

        # Tab Widget container
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Tab 1: Configuration
        self.tab_config = TabConfig()
        self.tabs.addTab(self.tab_config, "1. Configuration (Setup)")

        # Tab 2: Survey & Input Matrix
        self.tab_survey = TabSurvey()
        self.tabs.addTab(self.tab_survey, "2. Survey Matrix (Evaluation)")
        self.tabs.setTabEnabled(1, False)  # Khóa Tab 2 cho đến khi bấm START

        # Khi bấm START từ Tab 1, mở khóa và chuyển sang Tab 2
        self.tab_config.started.connect(self.on_survey_started)

    def on_survey_started(self, config):
        self.tab_survey.load_project(config)
        self.tabs.setTabEnabled(1, True)
        self.tabs.setCurrentIndex(1)