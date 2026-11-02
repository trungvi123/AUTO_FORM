
import os
import win32com.client

# Tạo thư mục kiểm tra
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_FILE = os.path.join(BASE_DIR, "test_word_edit.docx")

word = None
doc = None

try:
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = True

    doc = word.Documents.Add()

    selection = word.Selection
    selection.TypeText("AUTO FORM - KIEM TRA TICH HOP WORD")
    selection.TypeParagraph()
    selection.TypeText("Day la van ban kiem tra chuc nang chinh sua.")

    doc.SaveAs2(os.path.abspath(TEST_FILE))

    print("Tao tai lieu thanh cong!")
    print("Duong dan:", TEST_FILE)

except Exception as e:
    print("Loi:", e)

finally:
    # Giữ Word mở để bạn kiểm tra tài liệu
    pass
