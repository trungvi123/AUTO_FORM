
import sys
import os
import win32com.client
import shutil
import re

from docx import Document as WordDocument
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

from field_labels import FIELD_LABELS
from PySide6.QtWidgets import QGroupBox
from datetime import datetime
from PySide6.QtWidgets import QDialog
from PySide6.QtGui import QDesktopServices
from PySide6.QtCore import QUrl
from PySide6.QtCore import Qt,QTimer
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QMessageBox,
    QFormLayout, QListWidget, QScrollArea,QSizePolicy
)

from docxtpl import DocxTemplate

import database

# Kiểm tra máy có hỗ trợ xem PDF ngay trong ứng dụng không
try:
    from PySide6.QtPdf import QPdfDocument
    from PySide6.QtPdfWidgets import QPdfView
    PDF_VIEW_AVAILABLE = True
except ImportError:
    PDF_VIEW_AVAILABLE = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")


class AutoForm(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("AUTO FORM - Tạo biểu mẫu tự động")
        self.resize(950, 800)
        self.word_app = None
        self.word_doc = None
        self.editing_file = None

        self.edited_files = {}
        self.template_paths = []
        self.fields = {}
        self.current_profile_id = None
        # Bộ hẹn giờ tự động lưu
        
        database.create_database()

        self.init_ui()
        self.refresh_profiles()

    def init_ui(self):

        main_layout = QVBoxLayout()

        title = QLabel("AUTO FORM")
        title.setStyleSheet(
            "font-size: 28px; font-weight: bold; color: #2463a6;"
        )

        subtitle = QLabel(
            "Nhập dữ liệu, quản lý hồ sơ và tạo nhiều biểu mẫu Word tự động"
        )

        btn_refresh = QPushButton("REFRESH")
        btn_refresh.clicked.connect(self.refresh_form)

        # QUẢN LÝ HỒ SƠ
        profile_title = QLabel("QUẢN LÝ HỒ SƠ CÁ NHÂN")
        profile_title.setStyleSheet(
            "font-size: 18px; font-weight: bold;"
        )

        self.profile_search = QLineEdit()
        self.profile_search.setPlaceholderText(
            "Tìm kiếm theo họ tên hoặc số CCCD..."
        )
        self.profile_search.textChanged.connect(self.refresh_profiles)

        self.profile_list = QListWidget()
        self.profile_list.setMaximumHeight(85)

        btn_load = QPushButton("SỬ DỤNG HỒ SƠ")
        btn_load.clicked.connect(self.load_selected_profile)

        btn_save = QPushButton("LƯU / CẬP NHẬT HỒ SƠ")
        btn_save.clicked.connect(self.save_current_profile)

        btn_new = QPushButton("TẠO HỒ SƠ MỚI")
        btn_new.clicked.connect(self.create_new_profile)

        btn_delete = QPushButton("XÓA HỒ SƠ")
        btn_delete.clicked.connect(self.delete_selected_profile)

        profile_buttons = QHBoxLayout()
        profile_buttons.addWidget(btn_load)
        profile_buttons.addWidget(btn_new)
        profile_buttons.addWidget(btn_save)
        profile_buttons.addWidget(btn_delete)
        profile_buttons.addWidget(btn_refresh)

        # CHỌN BIỂU MẪU
        btn_select = QPushButton("CHỌN NHIỀU BIỂU MẪU WORD")
        btn_select.clicked.connect(self.select_templates)

        self.file_list = QListWidget()
        self.file_list.setMinimumHeight(75)

        self.file_label = QLabel("Chưa chọn biểu mẫu nào")

        # NHẬP THÔNG TIN
        form_title = QLabel("THÔNG TIN CẦN NHẬP")
        form_title.setStyleSheet(
            "font-size: 18px; font-weight: bold;"
        )

    

        self.form_layout = QFormLayout()

        self.form_widget = QWidget()
        self.form_widget.setLayout(self.form_layout)

        # Cho phép khu vực nhập liệu mở rộng theo nội dung
        self.form_widget.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Minimum
        )

        # XUẤT WORD
        btn_generate = QPushButton("TẠO TẤT CẢ BIỂU MẪU WORD")
        self.btn_edit_word = QPushButton("Chỉnh sửa bằng Microsoft Word")
        self.btn_finish_edit = QPushButton("Hoàn tất chỉnh sửa")

        self.btn_edit_word.clicked.connect(self.edit_selected_template)
        self.btn_finish_edit.clicked.connect(self.finish_word_edit)

        self.btn_finish_edit.setEnabled(False)

        main_layout.addWidget(self.btn_edit_word)
        main_layout.addWidget(self.btn_finish_edit)
        btn_generate.setMinimumHeight(45)

        btn_generate.setStyleSheet("""
            QPushButton {
                background-color: #2463a6;
                color: white;
                font-weight: bold;
                font-size: 15px;
                border-radius: 6px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #174b82;
            }
        """)

        btn_generate.clicked.connect(self.generate_word)

       # SẮP XẾP TOÀN BỘ GIAO DIỆN
        main_layout.addWidget(title)
        main_layout.addWidget(subtitle)
        main_layout.addSpacing(10)

        main_layout.addWidget(profile_title)
        main_layout.addWidget(self.profile_search)
        main_layout.addWidget(self.profile_list)
        main_layout.addLayout(profile_buttons)

        main_layout.addSpacing(15)

        main_layout.addWidget(btn_select)
        main_layout.addWidget(self.file_label)
        main_layout.addWidget(self.file_list)
        # Nút xem trước biểu mẫu
        btn_preview = QPushButton("XEM TRƯỚC BIỂU MẪU ĐANG CHỌN")
        btn_preview.setMinimumHeight(40)
        btn_preview.clicked.connect(self.preview_selected_template)

        self.file_list.itemDoubleClicked.connect(
            self.preview_selected_template
        )

        main_layout.addWidget(btn_preview)
        main_layout.addWidget(form_title)
        # Đưa toàn bộ các ô nhập liệu vào giao diện
        main_layout.addWidget(self.form_widget)

        main_layout.addSpacing(15)
        main_layout.addWidget(btn_generate)

        # Tạo nội dung toàn bộ ứng dụng
        container = QWidget()
        container.setLayout(main_layout)

        # Tạo thanh cuộn cho toàn bộ ứng dụng
        self.main_scroll = QScrollArea()
        self.main_scroll.setWidgetResizable(True)

        self.main_scroll.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )

        self.main_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )

        self.main_scroll.setWidget(container)

        # Đặt thanh cuộn làm giao diện chính
        self.setCentralWidget(self.main_scroll)

    # XÓA CÁC Ô NHẬP LIỆU CŨ
    def clear_form(self):

        while self.form_layout.count():
            item = self.form_layout.takeAt(0)

            if item.widget():
                item.widget().deleteLater()

        self.fields = {}

    def edit_selected_template(self):

        if not self.template_paths:
            QMessageBox.warning(
                self,
                "Thông báo",
                "Bạn chưa chọn mẫu Word!"
            )
            return

        selected_item = self.file_list.currentItem()

        if selected_item is None:
            QMessageBox.warning(
                self,
                "Thông báo",
                "Hãy chọn một mẫu Word!"
            )
            return

        filename = selected_item.text()

        template_path = next(
            (
                p for p in self.template_paths
                if os.path.basename(p) == filename
            ),
            None
        )

        if not template_path:
            QMessageBox.warning(
                self,
                "Lỗi",
                "Không tìm thấy mẫu Word!"
            )
            return

        try:
            data = self.collect_data()

            edit_dir = os.path.join(BASE_DIR, "edited")
            os.makedirs(edit_dir, exist_ok=True)

            edited_path = os.path.join(
                edit_dir,
                "editing_" + os.path.basename(template_path)
            )

            doc = DocxTemplate(template_path)
            doc.render(data)
            doc.save(edited_path)

            self.word_app = win32com.client.DispatchEx(
                "Word.Application"
            )

            self.word_app.Visible = True

            self.word_doc = self.word_app.Documents.Open(
                os.path.abspath(edited_path)
            )

            self.editing_file = edited_path
            self.editing_template_name = os.path.basename(template_path)
            self.btn_edit_word.setEnabled(False)
            self.btn_finish_edit.setEnabled(True)

            QMessageBox.information(
                self,
                "Thông báo",
                "Tài liệu đã mở trong Microsoft Word.\n"
                "Hãy chỉnh sửa nội dung rồi quay lại AUTO FORM."
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Lỗi",
                f"Không thể mở Microsoft Word:\n{str(e)}"
            )

    # HOÀN TẤT CHỈNH SỬA BẰNG MICROSOFT WORD
    def finish_word_edit(self):

        # Kiểm tra có tài liệu đang được chỉnh sửa không
        if self.word_doc is None or not self.editing_file:
            QMessageBox.warning(
                self,
                "Thông báo",
                "Không có tài liệu Word nào đang được chỉnh sửa!"
            )
            return

        try:
            # Lưu những thay đổi trong Microsoft Word
            self.word_doc.Save()

            # Đóng tài liệu Word
            self.word_doc.Close(False)

            # Thoát Microsoft Word
            self.word_app.Quit()

            # Ghi nhận file đã chỉnh sửa
            self.edited_files[self.editing_template_name] = (
                self.editing_file
            )

            # Xóa trạng thái tài liệu đang chỉnh sửa
            self.word_doc = None
            self.word_app = None
            self.editing_file = None
            self.editing_template_name = None

            # Cho phép chỉnh sửa biểu mẫu khác
            self.btn_edit_word.setEnabled(True)
            self.btn_finish_edit.setEnabled(False)

            QMessageBox.information(
                self,
                "Thành công",
                "Đã lưu và hoàn tất chỉnh sửa biểu mẫu Word!\n"
                "Khi tạo biểu mẫu, chương trình sẽ sử dụng "
                "bản đã chỉnh sửa."
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Lỗi hoàn tất chỉnh sửa",
                f"Không thể lưu hoặc đóng tài liệu Word:\n{str(e)}"
            )


    def get_template_variables_in_order(self, path):

        document = WordDocument(path)

        variables = []
        seen = set()

        pattern = re.compile(
            r"{{\s*([A-Za-z_][A-Za-z0-9_]*)\s*}}"
        )

        # Duyệt các đoạn văn theo đúng thứ tự trong tài liệu,
        # bao gồm cả những đoạn nằm trong bảng
        for paragraph_element in document.element.body.iter(qn("w:p")):

            paragraph = Paragraph(
                paragraph_element,
                document
            )

            matches = pattern.finditer(paragraph.text)

            for match in matches:

                variable = match.group(1)

                if variable not in seen:
                    variables.append(variable)
                    seen.add(variable)

        return variables
    
    def select_templates(self):

        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Chọn các biểu mẫu Word",
            "",
            "Word Files (*.docx)"
        )

        if not paths:
            return

        self.template_paths = paths
        self.file_list.clear()

        all_variables = []
        seen_variables = set()
        errors = []

        # Lưu dữ liệu hiện tại trước khi thay đổi các trường
        current_data = self.collect_data()

        for path in self.template_paths:
            try:
                doc = DocxTemplate(path)

                # Lấy danh sách biến mà DocxTemplate nhận diện
                declared_variables = (
                    doc.get_undeclared_template_variables()
                )

                # Lấy thứ tự biến xuất hiện trong Word
                ordered_variables = (
                    self.get_template_variables_in_order(path)
                )

                for field in ordered_variables:

                    if (
                        field in declared_variables
                        and field not in seen_variables
                    ):
                        all_variables.append(field)
                        seen_variables.add(field)

                self.file_list.addItem(
                    os.path.basename(path)
                )

            except Exception as e:
                errors.append(
                    f"{os.path.basename(path)}: {str(e)}"
                )

        # Xóa các trường cũ trên giao diện
        self.clear_form()

        # Tạo lại các trường theo đúng thứ tự trong Word
        for field in all_variables:

            input_box = QLineEdit()
            input_box.setPlaceholderText(
                f"Nhập {FIELD_LABELS.get(field, field)}"
            )

            # Giữ lại dữ liệu nếu trường đã tồn tại trước đó
            input_box.setText(
                str(current_data.get(field, "") or "")
            )

            self.fields[field] = input_box

            # Hiển thị tên tiếng Việt thay cho tên biến
            display_name = FIELD_LABELS.get(field, field)

            self.form_layout.addRow(
                display_name,
                input_box
            )

        self.file_label.setText(
            f"Đã chọn {len(self.template_paths)} biểu mẫu"
        )

        if errors:
            QMessageBox.warning(
                self,
                "Một số biểu mẫu gặp lỗi",
                "\n\n".join(errors)
            )

        if not all_variables and not errors:
            QMessageBox.information(
                self,
                "Thông báo",
                "Không tìm thấy trường thông tin nào trong biểu mẫu."
            )



    # LẤY DỮ LIỆU ĐANG NHẬP
    def collect_data(self):

        return {
            key: widget.text()
            for key, widget in self.fields.items()
        }

    # CẬP NHẬT DANH SÁCH HỒ SƠ
    def refresh_profiles(self):

        keyword = self.profile_search.text().strip()

        self.profile_list.clear()

        profiles = database.get_profiles(keyword)

        for profile in profiles:

            profile_id, full_name, identity_number, birth_date = profile

            display_text = (
                f"{full_name} | CCCD: {identity_number or 'Chưa có'}"
            )

            item = self.profile_list.addItem(display_text)
            list_item = self.profile_list.item(
                self.profile_list.count() - 1
            )

            list_item.setData(Qt.UserRole, profile_id)

    # LẤY ID HỒ SƠ ĐANG CHỌN
    def get_selected_profile_id(self):

        item = self.profile_list.currentItem()

        if item is None:
            return None

        return item.data(Qt.UserRole)

    # TẢI HỒ SƠ VÀO CÁC Ô NHẬP
    def load_selected_profile(self):

        profile_id = self.get_selected_profile_id()

        if profile_id is None:
            QMessageBox.warning(
                self,
                "Thông báo",
                "Bạn chưa chọn hồ sơ!"
            )
            return

        profile_data = database.get_profile(profile_id)

        if not profile_data:
            QMessageBox.warning(
                self,
                "Thông báo",
                "Không tìm thấy hồ sơ."
            )
            return

        self.current_profile_id = profile_id

        if not hasattr(self, "fields"):
            self.fields = {}

        # Hiển thị tất cả trường thông tin trong hồ sơ
        for key, value in profile_data.items():

            # Tạo trường nếu chưa tồn tại
            if key not in self.fields:

                input_box = QLineEdit()
                input_box.setPlaceholderText(f"Nhập {key}")

                self.fields[key] = input_box
                display_name = FIELD_LABELS.get(key, key)

                self.form_layout.addRow(
                    display_name,
                    input_box
                )

               

            # Tạm khóa tín hiệu khi điền dữ liệu hồ sơ
            self.fields[key].blockSignals(True)

            self.fields[key].setText(str(value or ""))

            # Bật lại tín hiệu
            self.fields[key].blockSignals(False)

        QMessageBox.information(
            self,
            "Thành công",
            "Đã tải thông tin hồ sơ vào biểu mẫu."
        )

    def create_new_profile(self):

        if not self.fields:
            QMessageBox.warning(
                self,
                "Thông báo",
                "Không có dữ liệu để tạo hồ sơ mới!"
            )
            return

        # Lấy toàn bộ dữ liệu đang hiển thị
        data = self.collect_data()

        # Kiểm tra họ tên
        full_name = (
            data.get("HO_TEN")
            or data.get("TEN")
            or data.get("HOTEN")
        )

        if not full_name:
            QMessageBox.warning(
                self,
                "Thiếu thông tin",
                "Bạn cần nhập họ tên trước khi tạo hồ sơ mới."
            )
            return

        # Chuẩn hóa họ tên
        data["HO_TEN"] = full_name

        try:
            # Luôn tạo hồ sơ mới, không cập nhật hồ sơ cũ
            profile_id = database.save_profile(
                data,
                profile_id=None
            )

            # Chuyển sang quản lý hồ sơ mới vừa tạo
            self.current_profile_id = profile_id

            # Cập nhật danh sách hồ sơ
            self.refresh_profiles()

            QMessageBox.information(
                self,
                "Thành công",
                f"Đã tạo hồ sơ mới cho {full_name} thành công!"
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Lỗi tạo hồ sơ",
                str(e)
            )
    
    # LƯU HOẶC CẬP NHẬT HỒ SƠ
    def save_current_profile(self):

        if not self.fields:
            QMessageBox.warning(
                self,
                "Thông báo",
                "Bạn cần chọn biểu mẫu và nhập thông tin trước."
            )
            return

        data = self.collect_data()

        # Bổ sung tên hồ sơ từ dữ liệu đang nhập
        full_name = (
            data.get("HO_TEN")
            or data.get("TEN")
            or data.get("HOTEN")
        )

        if not full_name:
            QMessageBox.warning(
                self,
                "Thiếu thông tin",
                "Bạn cần có trường HO_TEN và nhập họ tên trước khi lưu hồ sơ."
            )
            return

        # Giữ lại dữ liệu cũ nếu biểu mẫu hiện tại không có trường đó
        if self.current_profile_id:
            old_data = database.get_profile(self.current_profile_id) or {}
            old_data.update(data)
            data = old_data

        # Chuẩn hóa trường họ tên để lưu vào cơ sở dữ liệu
        data["HO_TEN"] = full_name

        try:
            profile_id = database.save_profile(
                data,
                self.current_profile_id
            )

            self.current_profile_id = profile_id
            self.refresh_profiles()

            QMessageBox.information(
                self,
                "Thành công",
                "Đã lưu hồ sơ thành công!"
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Lỗi lưu hồ sơ",
                str(e)
            )
    # refresh_form
    def refresh_form(self):

        # Xác nhận trước khi xóa dữ liệu
        reply = QMessageBox.question(
            self,
            "Xác nhận làm mới",
            "Bạn có chắc chắn muốn xóa toàn bộ dữ liệu đang nhập không?",
            QMessageBox.StandardButton.Yes |
            QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        # Xóa nội dung nhưng giữ nguyên các trường
        for widget in self.fields.values():
            widget.blockSignals(True)
            widget.clear()
            widget.blockSignals(False)

        # Chuyển sang trạng thái tạo hồ sơ mới
        self.current_profile_id = None

        QMessageBox.information(
            self,
            "Thông báo",
            "Đã làm mới các trường dữ liệu!"
        )   
   
    # XÓA HỒ SƠ
    def delete_selected_profile(self):

        profile_id = self.get_selected_profile_id()

        if profile_id is None:
            QMessageBox.warning(
                self,
                "Thông báo",
                "Bạn chưa chọn hồ sơ cần xóa."
            )
            return

        answer = QMessageBox.question(
            self,
            "Xác nhận",
            "Bạn có chắc chắn muốn xóa hồ sơ này?"
        )

        if answer == QMessageBox.StandardButton.Yes:

            database.delete_profile(profile_id)

            if self.current_profile_id == profile_id:
                self.current_profile_id = None

            self.refresh_profiles()

            QMessageBox.information(
                self,
                "Thành công",
                "Đã xóa hồ sơ."
            )


    # XEM TRƯỚC BIỂU MẪU ĐANG CHỌN
    def preview_selected_template(self):

        # Kiểm tra đã chọn biểu mẫu chưa
        if not self.template_paths:
            QMessageBox.warning(
                self,
                "Thông báo",
                "Bạn chưa chọn biểu mẫu Word!"
            )
            return

        # Lấy biểu mẫu đang được chọn trong danh sách
        selected_item = self.file_list.currentItem()

        if selected_item is None:
            QMessageBox.warning(
                self,
                "Thông báo",
                "Bạn hãy chọn một biểu mẫu trong danh sách!"
            )
            return

        filename = selected_item.text()

        # Tìm đường dẫn tương ứng
        template_path = None

        for path in self.template_paths:
            if os.path.basename(path) == filename:
                template_path = path
                break

        if not template_path:
            QMessageBox.warning(
                self,
                "Lỗi",
                "Không tìm thấy đường dẫn biểu mẫu."
            )
            return

        try:

            # Thu thập dữ liệu hiện tại
            data = self.collect_data()

            # Tạo thư mục xem trước
            preview_dir = os.path.join(BASE_DIR, "preview")
            os.makedirs(preview_dir, exist_ok=True)

            # Tạo tên file tạm thời, tránh trùng khi xem nhiều lần
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

            preview_docx = os.path.join(
                preview_dir,
                f"preview_{timestamp}.docx"
            )

            preview_pdf = os.path.join(
                preview_dir,
                f"preview_{timestamp}.pdf"
            )

            # Điền dữ liệu vào bản sao biểu mẫu
            doc = DocxTemplate(template_path)
            doc.render(data)
            doc.save(preview_docx)

            # Dùng Microsoft Word chuyển DOCX sang PDF
            import win32com.client

            word = win32com.client.DispatchEx(
                "Word.Application"
            )

            word.Visible = False
            word.DisplayAlerts = 0

            word_doc = None

            try:

                word_doc = word.Documents.Open(
                    os.path.abspath(preview_docx)
                )

                # 17 là định dạng PDF của Microsoft Word
                word_doc.SaveAs(
                    os.path.abspath(preview_pdf),
                    FileFormat=17
                )

            finally:

                if word_doc is not None:
                    word_doc.Close(False)

                word.Quit()

            if not os.path.isfile(preview_pdf):
                raise Exception(
                    "Microsoft Word không tạo được file PDF."
                )

            # Hiển thị PDF ngay trong AUTO FORM nếu được hỗ trợ
            if PDF_VIEW_AVAILABLE:

                self.show_pdf_preview(preview_pdf)

            else:

                # Nếu không hỗ trợ Qt PDF thì mở bằng ứng dụng mặc định
                QDesktopServices.openUrl(
                    QUrl.fromLocalFile(preview_pdf)
                )

                QMessageBox.information(
                    self,
                    "Xem trước",
                    "Đã tạo bản xem trước PDF và mở bằng trình đọc PDF."
                )

        except Exception as e:

            QMessageBox.critical(
                self,
                "Lỗi xem trước",
                f"Không thể xem trước biểu mẫu.\n\nChi tiết:\n{str(e)}"
            )


    # HIỂN THỊ PDF TRONG CỬA SỔ AUTO FORM
    def show_pdf_preview(self, pdf_path):

        dialog = QDialog(self)

        dialog.setWindowTitle("AUTO FORM - Xem trước biểu mẫu")
        dialog.resize(950, 750)

        layout = QVBoxLayout(dialog)

        pdf_view = QPdfView(dialog)
        pdf_document = QPdfDocument(dialog)

        result = pdf_document.load(pdf_path)

        if result != QPdfDocument.Error.None_:
            QMessageBox.warning(
                self,
                "Lỗi PDF",
                "Không thể đọc file PDF."
            )
            return

        pdf_view.setDocument(pdf_document)

        pdf_view.setPageMode(
            QPdfView.PageMode.MultiPage
        )

        pdf_view.setZoomMode(
            QPdfView.ZoomMode.FitToWidth
        )

        layout.addWidget(pdf_view)

        btn_close = QPushButton("ĐÓNG BẢN XEM TRƯỚC")
        btn_close.clicked.connect(dialog.close)

        layout.addWidget(btn_close)

        # Giữ tài liệu PDF trong thời gian cửa sổ đang mở
        dialog.pdf_document = pdf_document
        dialog.pdf_view = pdf_view

        dialog.exec()
    
    # TẠO CÁC BIỂU MẪU WORD
    def generate_word(self):

        if not self.template_paths:
            QMessageBox.warning(
                self,
                "Thông báo",
                "Bạn chưa chọn biểu mẫu Word!"
            )
            return

        data = self.collect_data()

        os.makedirs(OUTPUT_DIR, exist_ok=True)

        success_count = 0
        errors = []
        edited_count = 0

        for index, path in enumerate(self.template_paths, start=1):

            try:
                filename = os.path.basename(path)

                output_name = (
                    f"{index:02d}_"
                    f"{os.path.splitext(filename)[0]}"
                    f"_completed.docx"
                )

                output_path = os.path.join(
                    OUTPUT_DIR,
                    output_name
                )

                # KIỂM TRA BIỂU MẪU ĐÃ ĐƯỢC CHỈNH SỬA HAY CHƯA
                if filename in self.edited_files:

                    edited_path = self.edited_files[filename]

                    if os.path.isfile(edited_path):

                        # Sử dụng file đã chỉnh sửa bằng Word
                        shutil.copy2(
                            edited_path,
                            output_path
                        )

                        edited_count += 1

                    else:
                        raise Exception(
                            "Không tìm thấy file đã chỉnh sửa!"
                        )

                else:

                    # Nếu chưa chỉnh sửa thì tạo từ biểu mẫu gốc
                    doc = DocxTemplate(path)
                    doc.render(data)
                    doc.save(output_path)

                success_count += 1

            except Exception as e:

                errors.append(
                    f"{os.path.basename(path)}: {str(e)}"
                )

        message = (
            f"Đã tạo thành công {success_count}/"
            f"{len(self.template_paths)} biểu mẫu.\n\n"
            f"Số biểu mẫu sử dụng bản đã chỉnh sửa: "
            f"{edited_count}\n\n"
            f"Thư mục lưu:\n{OUTPUT_DIR}"
        )

        if errors:
            message += "\n\nCác biểu mẫu gặp lỗi:\n"
            message += "\n".join(errors)

        QMessageBox.information(
            self,
            "Kết quả tạo biểu mẫu",
            message
        )


if __name__ == "__main__":

    app = QApplication(sys.argv)

    window = AutoForm()
    window.show()

    sys.exit(app.exec())
