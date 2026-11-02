
import sys

try:
    from PySide6.QtAxContainer import QAxWidget

    print("Qt ActiveX: Có thể import QAxWidget")

except ImportError as e:
    print("Qt ActiveX: Không khả dụng")
    print("Chi tiết:", e)

try:
    import win32com.client

    word = win32com.client.Dispatch("Word.Application")

    print("Microsoft Word: Kết nối thành công")
    print("Phiên bản:", word.Version)

    word.Quit()

except Exception as e:
    print("Microsoft Word: Không kết nối được")
    print("Chi tiết:", e)
