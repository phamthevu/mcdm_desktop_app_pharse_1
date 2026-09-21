from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QGroupBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QPushButton,
    QComboBox, QLineEdit, QMessageBox, QFileDialog, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QBrush, QFont
from core.models import ProjectConfig
from core.excel_handler import export_all_experts_results
from ui.widgets.z_cell_widget import MatrixCellWidget


class TabSurvey(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.config = ProjectConfig()
        self.row_labels = []
        self.col_labels = []
        self.cell_widgets = {}
        
        # Bộ nhớ lưu kết quả của TẤT CẢ chuyên gia:
        # { "Expert_Name": { "code_matrix": [...], "value_matrix": [...] } }
        self.saved_experts_data = {}
        self.current_expert_name = ""

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(12, 12, 12, 12)

        # =====================================================================
        # 1. THANH ĐIỀU HƯỚNG CHUYÊN GIA & NÚT LƯU
        # =====================================================================
        grp_session = QGroupBox("1. Expert Management & Switcher")
        lay_session = QHBoxLayout(grp_session)

        lay_session.addWidget(QLabel("Select Expert:"))
        self.cb_experts = QComboBox()
        self.cb_experts.setMinimumWidth(220)
        self.cb_experts.currentIndexChanged.connect(self.on_expert_changed)

        self.txt_expert_free = QLineEdit()
        self.txt_expert_free.setPlaceholderText("Enter Expert Name...")
        self.txt_expert_free.setVisible(False)

        lay_session.addWidget(self.cb_experts)
        lay_session.addWidget(self.txt_expert_free)

        # Nút Lưu kết quả chuyên gia hiện tại
        self.btn_save_expert = QPushButton("💾 Save Current Expert Data")
        self.btn_save_expert.setFixedHeight(34)
        self.btn_save_expert.setStyleSheet("""
            QPushButton {
                background-color: #0d6efd;
                color: white;
                font-weight: bold;
                border-radius: 4px;
                padding: 0 12px;
            }
            QPushButton:hover { background-color: #0b5ed7; }
        """)
        self.btn_save_expert.clicked.connect(self.save_current_expert)
        lay_session.addWidget(self.btn_save_expert)

        self.lbl_saved_status = QLabel("Status: [Not Saved Yet]")
        self.lbl_saved_status.setStyleSheet("font-weight: bold; color: #dc3545;")
        lay_session.addWidget(self.lbl_saved_status)

        lay_session.addStretch()

        self.lbl_model_badge = QLabel("Model: DEMATEL")
        self.lbl_model_badge.setStyleSheet("font-weight: bold; color: #0d6efd; padding: 4px 8px; border: 1px solid #0d6efd; border-radius: 4px;")
        lay_session.addWidget(self.lbl_model_badge)

        self.lbl_z_badge = QLabel("Z-Number: OFF")
        self.lbl_z_badge.setStyleSheet("font-weight: bold; color: #198754; padding: 4px 8px; border: 1px solid #198754; border-radius: 4px;")
        lay_session.addWidget(self.lbl_z_badge)

        main_layout.addWidget(grp_session)

        # =====================================================================
        # 2. KHUNG HƯỚNG DẪN ĐÁNH GIÁ (BANNER LỚN, RÕ RÀNG)
        # =====================================================================
        frame_guide = QFrame()
        frame_guide.setStyleSheet("""
            QFrame {
                background-color: #fff9db;
                border: 2px solid #fab005;
                border-radius: 6px;
                padding: 6px;
            }
        """)
        lay_guide = QVBoxLayout(frame_guide)

        self.lbl_guide_title = QLabel("📌 EVALUATION INSTRUCTION FOR SELECTED CELL:")
        self.lbl_guide_title.setStyleSheet("color: #d9480f; font-weight: bold; font-size: 13px;")
        lay_guide.addWidget(self.lbl_guide_title)

        self.lbl_dynamic_guide = QLabel("Click any cell in the table below to see the evaluation question.")
        self.lbl_dynamic_guide.setStyleSheet("font-size: 15px; font-weight: bold; color: #1c7ed6; margin: 4px 0;")
        self.lbl_dynamic_guide.setWordWrap(True)
        lay_guide.addWidget(self.lbl_dynamic_guide)

        self.lbl_scale_ref = QLabel("Scale Definition Reference:")
        self.lbl_scale_ref.setStyleSheet("color: #495057; font-size: 12px;")
        self.lbl_scale_ref.setWordWrap(True)
        lay_guide.addWidget(self.lbl_scale_ref)

        main_layout.addWidget(frame_guide)

        # =====================================================================
        # 3. MA TRẬN ĐÁNH GIÁ
        # =====================================================================
        self.table_matrix = QTableWidget()
        self.table_matrix.cellClicked.connect(self.on_cell_clicked)
        main_layout.addWidget(self.table_matrix)

        # =====================================================================
        # 4. THANH NÚT XUẤT TOÀN BỘ KẾT QUẢ
        # =====================================================================
        lay_footer = QHBoxLayout()
        self.lbl_total_saved_counter = QLabel("Experts completed: 0")
        self.lbl_total_saved_counter.setStyleSheet("font-weight: bold; font-size: 13px; color: #212529;")
        lay_footer.addWidget(self.lbl_total_saved_counter)

        lay_footer.addStretch()

        self.btn_export_all = QPushButton("📤 EXPORT ALL EXPERTS (2 Excel Files: Codes & Values)")
        self.btn_export_all.setFixedHeight(44)
        self.btn_export_all.setStyleSheet("""
            QPushButton {
                background-color: #198754;
                color: white;
                font-weight: bold;
                font-size: 14px;
                border-radius: 5px;
                padding: 0 20px;
            }
            QPushButton:hover { background-color: #157347; }
        """)
        self.btn_export_all.clicked.connect(self.on_export_all_clicked)
        lay_footer.addWidget(self.btn_export_all)

        main_layout.addLayout(lay_footer)

    def load_project(self, config: ProjectConfig):
        """Khởi tạo dự án từ cấu hình Config."""
        self.config = config
        self.saved_experts_data.clear()
        self.cell_widgets.clear()

        self.lbl_model_badge.setText(f"Model: {config.model_type}")
        self.lbl_z_badge.setText(f"Z-Number: {'ON' if config.use_z_number else 'OFF'}")

        # Nạp Chuyên gia
        self.cb_experts.blockSignals(True)
        self.cb_experts.clear()
        if config.expert_auth_mode == "Account":
            self.cb_experts.setVisible(True)
            self.txt_expert_free.setVisible(False)
            for exp in config.experts:
                self.cb_experts.addItem(f"{exp.name} ({exp.username})", userData=exp.name)
            self.current_expert_name = self.cb_experts.currentData() if config.experts else ""
        else:
            self.cb_experts.setVisible(False)
            self.txt_expert_free.setVisible(True)
            self.txt_expert_free.setText("Expert_01")
            self.current_expert_name = "Expert_01"
        self.cb_experts.blockSignals(False)

        self.update_scale_reference_text()
        self.build_matrix()
        self.update_saved_counter()

    def update_scale_reference_text(self):
        if self.config.scale_type == "Continuous":
            self.lbl_scale_ref.setText(f"Scale Reference: Continuous Slider [{self.config.continuous_min} -> {self.config.continuous_max}]")
        elif not self.config.use_z_number:
            items = [f"<b>{s.code}</b>: {s.linguistic} ({s.value})" for s in self.config.standard_scales]
            self.lbl_scale_ref.setText("Scale Reference: " + " | ".join(items))
        else:
            items_a = [f"<b>{s.code}</b>: {s.linguistic}" for s in self.config.z_scales_a]
            items_b = [f"<b>{s.code}</b>: {s.linguistic}" for s in self.config.z_scales_b]
            self.lbl_scale_ref.setText(f"<b>[Eval A]:</b> {' | '.join(items_a)}<br><b>[Certainty B]:</b> {' | '.join(items_b)}")

    def build_matrix(self):
        """Sinh bảng ma trận theo chuẩn mô hình."""
        model = self.config.model_type
        factors = [f.code for f in self.config.factors]
        alts = [a.code for a in self.config.alternatives]

        scales_a = self.config.z_scales_a if self.config.use_z_number else self.config.standard_scales
        scales_b = self.config.z_scales_b if self.config.use_z_number else []

        self.cell_widgets.clear()

        if model in ["DEMATEL", "AHP"]:
            self.row_labels = factors
            self.col_labels = factors
            n = len(factors)
            self.table_matrix.setRowCount(n)
            self.table_matrix.setColumnCount(n)
            self.table_matrix.setHorizontalHeaderLabels(factors)
            self.table_matrix.setVerticalHeaderLabels(factors)

            for i in range(n):
                for j in range(n):
                    if model == "DEMATEL" and i == j:
                        item = QTableWidgetItem("0")
                        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                        item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                        item.setBackground(QBrush(QColor("#dee2e6")))
                        self.table_matrix.setItem(i, j, item)
                    elif model == "AHP" and i == j:
                        item = QTableWidgetItem("1")
                        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                        item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                        item.setBackground(QBrush(QColor("#dee2e6")))
                        self.table_matrix.setItem(i, j, item)
                    elif model == "AHP" and i > j:
                        item = QTableWidgetItem("1 / (a_ji)")
                        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                        item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                        item.setBackground(QBrush(QColor("#fff3cd")))
                        self.table_matrix.setItem(i, j, item)
                    else:
                        cell = MatrixCellWidget(
                            is_continuous=(self.config.scale_type == "Continuous"),
                            use_z=self.config.use_z_number,
                            scales_a=scales_a, scales_b=scales_b,
                            c_min=self.config.continuous_min, c_max=self.config.continuous_max
                        )
                        self.table_matrix.setCellWidget(i, j, cell)
                        self.cell_widgets[(i, j)] = cell

        elif model == "TOPSIS":
            self.row_labels = alts
            self.col_labels = factors
            r_count = len(alts)
            c_count = len(factors)
            self.table_matrix.setRowCount(r_count)
            self.table_matrix.setColumnCount(c_count)
            self.table_matrix.setHorizontalHeaderLabels(factors)
            self.table_matrix.setVerticalHeaderLabels(alts)

            for i in range(r_count):
                for j in range(c_count):
                    cell = MatrixCellWidget(
                        is_continuous=(self.config.scale_type == "Continuous"),
                        use_z=self.config.use_z_number,
                        scales_a=scales_a, scales_b=scales_b,
                        c_min=self.config.continuous_min, c_max=self.config.continuous_max
                    )
                    self.table_matrix.setCellWidget(i, j, cell)
                    self.cell_widgets[(i, j)] = cell

        elif model == "DELPHI":
            self.row_labels = factors
            self.col_labels = ["Rating"]
            n = len(factors)
            self.table_matrix.setRowCount(n)
            self.table_matrix.setColumnCount(1)
            self.table_matrix.setHorizontalHeaderLabels(["Rating"])
            self.table_matrix.setVerticalHeaderLabels([f"{f.code}: {f.name}" for f in self.config.factors])

            for i in range(n):
                cell = MatrixCellWidget(
                    is_continuous=(self.config.scale_type == "Continuous"),
                    use_z=self.config.use_z_number,
                    scales_a=scales_a, scales_b=scales_b,
                    c_min=self.config.continuous_min, c_max=self.config.continuous_max
                )
                self.table_matrix.setCellWidget(i, 0, cell)
                self.cell_widgets[(i, 0)] = cell

        self.table_matrix.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_matrix.verticalHeader().setDefaultSectionSize(42)

    def on_cell_clicked(self, row: int, col: int):
        """Hiển thị câu hỏi hướng dẫn TO, RÕ RÀNG theo từng ngữ cảnh ô (i, j)."""
        model = self.config.model_type
        r_code = self.row_labels[row] if row < len(self.row_labels) else f"Row {row+1}"
        c_code = self.col_labels[col] if col < len(self.col_labels) else f"Col {col+1}"

        # Tìm tên chi tiết
        r_name = next((f.name for f in self.config.factors + self.config.alternatives if f.code == r_code), r_code)
        c_name = next((f.name for f in self.config.factors + self.config.alternatives if f.code == c_code), c_code)

        if model == "DEMATEL":
            if row == col:
                self.lbl_dynamic_guide.setText(f"Cell ({r_code}, {c_code}): A factor cannot influence itself. Locked to 0.")
            else:
                self.lbl_dynamic_guide.setText(
                    f"👉 How much direct influence does [{r_code} - {r_name}] exert on [{c_code} - {c_name}]?"
                )
        elif model == "AHP":
            if row == col:
                self.lbl_dynamic_guide.setText(f"Cell ({r_code}, {c_code}): Criteria compared to itself. Locked to 1.")
            elif row < col:
                self.lbl_dynamic_guide.setText(
                    f"👉 How much more important is [{r_code} - {r_name}] compared to [{c_code} - {c_name}]?"
                )
            else:
                self.lbl_dynamic_guide.setText(f"Cell ({r_code}, {c_code}): Reciprocal cell 1 / ({c_code}, {r_code}).")
        elif model == "TOPSIS":
            self.lbl_dynamic_guide.setText(
                f"👉 How does Alternative [{r_code} - {r_name}] perform with respect to Criterion [{c_code} - {c_name}]?"
            )
        elif model == "DELPHI":
            self.lbl_dynamic_guide.setText(f"👉 Rate the significance or consensus level for: [{r_code} - {r_name}].")

    # =========================================================================
    # LƯU TRỮ VÀ CHUYỂN ĐỔI CHUYÊN GIA
    # =========================================================================
    def get_current_matrices(self):
        """Thu thập ma trận Code và Value của giao diện hiện tại."""
        r_len = len(self.row_labels)
        c_len = len(self.col_labels)

        code_mat = [["" for _ in range(c_len)] for _ in range(r_len)]
        val_mat = [["" for _ in range(c_len)] for _ in range(r_len)]

        for r in range(r_len):
            for c in range(c_len):
                if (r, c) in self.cell_widgets:
                    widget = self.cell_widgets[(r, c)]
                    code_mat[r][c] = widget.get_code()
                    val_mat[r][c] = widget.get_value()
                else:
                    item = self.table_matrix.item(r, c)
                    txt = item.text() if item else ""
                    code_mat[r][c] = txt
                    val_mat[r][c] = txt

        return code_mat, val_mat

    def save_current_expert(self):
        """Lưu ma trận của chuyên gia đang chọn vào bộ nhớ."""
        exp_name = self.cb_experts.currentData() if self.config.expert_auth_mode == "Account" else self.txt_expert_free.text().strip()
        if not exp_name:
            QMessageBox.warning(self, "Warning", "Please enter/select a valid expert name.")
            return

        code_mat, val_mat = self.get_current_matrices()
        self.saved_experts_data[exp_name] = {
            "code_matrix": code_mat,
            "value_matrix": val_mat
        }

        self.lbl_saved_status.setText(f"Status: [Saved for {exp_name}]")
        self.lbl_saved_status.setStyleSheet("font-weight: bold; color: #198754;")
        self.update_saved_counter()
        QMessageBox.information(self, "Saved", f"Data for expert '{exp_name}' has been successfully saved!")

    def on_expert_changed(self):
        """Khi chuyển dropdown sang chuyên gia khác: Tự động load dữ liệu của họ nếu có."""
        new_exp = self.cb_experts.currentData() if self.config.expert_auth_mode == "Account" else self.txt_expert_free.text().strip()
        self.current_expert_name = new_exp

        if new_exp in self.saved_experts_data:
            # Khôi phục dữ liệu đã lưu của người này
            saved = self.saved_experts_data[new_exp]["code_matrix"]
            for (r, c), widget in self.cell_widgets.items():
                if r < len(saved) and c < len(saved[r]):
                    widget.set_value_by_code(saved[r][c])
            self.lbl_saved_status.setText(f"Status: [Loaded saved data]")
            self.lbl_saved_status.setStyleSheet("font-weight: bold; color: #198754;")
        else:
            self.lbl_saved_status.setText("Status: [Not Saved Yet]")
            self.lbl_saved_status.setStyleSheet("font-weight: bold; color: #dc3545;")

    def update_saved_counter(self):
        total_experts = len(self.config.experts) if self.config.expert_auth_mode == "Account" else len(self.saved_experts_data)
        saved_count = len(self.saved_experts_data)
        self.lbl_total_saved_counter.setText(f"Experts Completed: {saved_count} / {max(total_experts, saved_count)}")

    # =========================================================================
    # XUẤT TOÀN BỘ CHUYÊN GIA
    # =========================================================================
    def on_export_all_clicked(self):
        # Tự động lưu người hiện tại nếu chưa lưu
        cur_name = self.cb_experts.currentData() if self.config.expert_auth_mode == "Account" else self.txt_expert_free.text().strip()
        if cur_name and cur_name not in self.saved_experts_data:
            code_mat, val_mat = self.get_current_matrices()
            self.saved_experts_data[cur_name] = {"code_matrix": code_mat, "value_matrix": val_mat}
            self.update_saved_counter()

        if not self.saved_experts_data:
            QMessageBox.warning(self, "No Data", "No expert data has been saved yet to export.")
            return

        folder = QFileDialog.getExistingDirectory(self, "Select Folder to Save 2 Result Files")
        if not folder:
            return

        file_code = f"{folder}/ALL_EXPERTS_{self.config.model_type}_RAW_CODES.xlsx"
        file_val = f"{folder}/ALL_EXPERTS_{self.config.model_type}_NUMERICAL_VALUES.xlsx"

        try:
            export_all_experts_results(
                file_code, file_val, self.config,
                self.saved_experts_data, self.row_labels, self.col_labels
            )
            QMessageBox.information(
                self, "Export Completed Successfully!",
                f"Successfully exported data of {len(self.saved_experts_data)} experts to:\n\n"
                f"1. Raw Codes File:\n{file_code}\n\n"
                f"2. Numerical/Fuzzy Values File:\n{file_val}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"An error occurred while exporting:\n{e}")