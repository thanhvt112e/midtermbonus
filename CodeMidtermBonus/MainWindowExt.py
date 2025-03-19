import os
import webbrowser
import pandas as pd
import plotly.express as px
from tempfile import NamedTemporaryFile
from PyQt6.QtWidgets import QFileDialog, QMessageBox

from MainWindow import Ui_MainWindow


class MainWindowExt(Ui_MainWindow):
    def __init__(self, main_window):
        super().setupUi(main_window)
        self.main_window = main_window

        # Khởi tạo dữ liệu
        self.excel_data = None
        self.fig = None
        self.temp_html_file = None

        # Kết nối sự kiện cho các nút bấm
        self.pushButtonPickFile.clicked.connect(self.browse_file)
        self.pushButtonOpen.clicked.connect(self.open_chart_in_browser)
        self.pushButtonSave.clicked.connect(self.save_chart_to_html)


    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            None, "Chọn file Excel", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.lineEditFile.setText(file_path)
            try:
                self.process_excel_data(file_path)
                QMessageBox.information(None, "Thành công", "Dữ liệu đã được xử lý thành công!")
                self.pushButtonOpen.setEnabled(True)
                self.pushButtonSave.setEnabled(True)
            except Exception as e:
                QMessageBox.critical(None, "Lỗi", f"Lỗi: {str(e)}")
                self.pushButtonOpen.setEnabled(False)
                self.pushButtonSave.setEnabled(False)

    def process_excel_data(self, file_path):
        """Xử lý file Excel và tạo biểu đồ"""
        try:
            self.excel_data = pd.read_excel(file_path)
            if 'Học kỳ' not in self.excel_data.columns:
                # Tìm cột có thể là học kỳ
                semester_cols = [col for col in self.excel_data.columns if
                                 any(kw in str(col).lower() for kw in ['kỳ', 'ky', 'hk', 'semester'])]
                if semester_cols:
                    self.excel_data.rename(columns={semester_cols[0]: 'Học kỳ'}, inplace=True)
                else:
                    # Tạo cột học kỳ mặc định
                    self.excel_data['Học kỳ'] = 1

            if 'Loại' not in self.excel_data.columns:
                # Tìm cột có thể là loại
                type_cols = [col for col in self.excel_data.columns if
                             any(kw in str(col).lower() for kw in ['loại', 'loai', 'type', 'bắt buộc', 'tự chọn'])]
                if type_cols:
                    self.excel_data.rename(columns={type_cols[0]: 'Loại'}, inplace=True)
                else:
                    # Tạo cột loại mặc định
                    self.excel_data['Loại'] = 'Bắt buộc'

            if 'Tên môn học' not in self.excel_data.columns:
                # Tìm cột có thể là tên môn học
                name_cols = [col for col in self.excel_data.columns if
                             any(kw in str(col).lower() for kw in ['tên', 'ten', 'môn', 'mon', 'name', 'course'])]
                if name_cols:
                    self.excel_data.rename(columns={name_cols[0]: 'Tên môn học'}, inplace=True)
                else:
                    # Tìm cột kiểu chuỗi đầu tiên không phải Loại
                    string_cols = [col for col in self.excel_data.columns if
                                   self.excel_data[col].dtype == 'object' and col != 'Loại']
                    if string_cols:
                        self.excel_data.rename(columns={string_cols[0]: 'Tên môn học'}, inplace=True)
                    else:
                        # Tạo cột tên môn học mặc định
                        self.excel_data['Tên môn học'] = [f'Môn học {i + 1}' for i in range(len(self.excel_data))]

            if 'Số tín chỉ' not in self.excel_data.columns:
                # Tìm cột có thể là số tín chỉ
                credit_cols = [col for col in self.excel_data.columns if
                               any(kw in str(col).lower() for kw in ['tín', 'tin', 'tc', 'credit'])]
                if credit_cols:
                    self.excel_data.rename(columns={credit_cols[0]: 'Số tín chỉ'}, inplace=True)
                else:
                    # Tìm cột số đầu tiên không phải Học kỳ
                    numeric_cols = [col for col in self.excel_data.select_dtypes(include=['number']).columns if
                                    col != 'Học kỳ']
                    if numeric_cols:
                        print(f"Sử dụng cột '{numeric_cols[0]}' làm 'Số tín chỉ'")
                        self.excel_data.rename(columns={numeric_cols[0]: 'Số tín chỉ'}, inplace=True)
                    else:
                        # Tạo cột số tín chỉ mặc định
                        self.excel_data['Số tín chỉ'] = 3
            # Chuyển đổi kiểu dữ liệu
            self.excel_data['Học kỳ'] = pd.to_numeric(self.excel_data['Học kỳ'], errors='coerce').fillna(
                1).astype(int)
            self.excel_data['Số tín chỉ'] = pd.to_numeric(self.excel_data['Số tín chỉ'],
                                                          errors='coerce').fillna(
                3).astype(int)

            # Chuẩn hóa cột Loại
            def standardize_type(type_str):
                if pd.isna(type_str):
                    return "Bắt buộc"

                type_lower = str(type_str).lower()
                if any(kw in type_lower for kw in ["bắt buộc", "bat buoc", "bb", "bắt", "bat"]):
                    return "Bắt buộc"
                elif any(kw in type_lower for kw in
                         ["tự chọn", "tu chon", "tc", "tự", "tu", "chọn", "chon"]):
                    return "Tự chọn"
                return "Bắt buộc"  # Mặc định là bắt buộc

            self.excel_data['Loại'] = self.excel_data['Loại'].apply(standardize_type)

            print(self.excel_data.head())
            self.excel_data['path'] = self.excel_data.apply(
                lambda
                    row: f"Chương trình đào tạo/Học kỳ {row['Học kỳ']}/{row['Loại']}/{row['Tên môn học']}",
                axis=1
            )

            self.fig = px.sunburst(
                self.excel_data,
                path=['Học kỳ', 'Loại', 'Tên môn học'],
                values='Số tín chỉ',
                title="Chương trình đào tạo Thương Mại Điện Tử",
                width=1000,
                height=1000,
                color='Học kỳ',
                color_continuous_scale='viridis'
            )

            # Cấu hình layout
            self.fig.update_layout(margin=dict(t=30, l=0, r=0, b=0))

        except Exception as e:
            import traceback
            traceback.print_exc()
            raise ValueError(f"Lỗi khi xử lý file Excel: {str(e)}")

    def open_chart_in_browser(self):
        if self.fig:
            try:
                # Tạo file HTML mới
                self.temp_html_file = NamedTemporaryFile(delete=False, suffix='.html')
                temp_path = self.temp_html_file.name
                self.temp_html_file.close()

                self.fig.write_html(temp_path, full_html=True, include_plotlyjs='cdn')
                webbrowser.open('file://' + temp_path)
                QMessageBox.information(None, "Thành công", "Đã mở biểu đồ trong trình duyệt!")
            except Exception as e:
                import traceback
                traceback.print_exc()
                QMessageBox.critical(None, "Lỗi", f"Lỗi khi mở biểu đồ: {str(e)}")

    def save_chart_to_html(self):
        if self.fig:
            try:
                file_path, _ = QFileDialog.getSaveFileName(
                    None, "Lưu biểu đồ", "411_10k.html", "HTML Files (*.html)"
                )

                if file_path:
                    self.fig.write_html(file_path, full_html=True, include_plotlyjs='cdn')
                    webbrowser.open('file://' + file_path)
                    QMessageBox.information(None, "Thành công", f"Đã lưu biểu đồ tại: {file_path}")
            except Exception as e:
                import traceback
                traceback.print_exc()
                QMessageBox.critical(None, "Lỗi", f"Lỗi khi lưu biểu đồ: {str(e)}")

    def closeEvent(self, event):
        if self.temp_html_file:
            try:
                os.unlink(self.temp_html_file.name)
            except:
                pass
        event.accept()