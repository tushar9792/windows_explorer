# import tkinter as tk
# from tkinter import filedialog
# import os
#
# def open_file():
#     root = tk.Tk()
#     root.withdraw()  # Hide the main window
#
#     file_path = filedialog.askopenfilename(
#         initialdir="C:\\Users\\YourUserName\\Documents",  # Replace with your desired initial directory
#         title="Select a .doc file",
#         filetypes=[("Word Documents, Excel Documents", "*.doc;*.docx;*.xls")]
#     )
#
#     if file_path:
#         os.startfile(file_path)
#
# open_file()


# import tkinter as tk
# from tkinter import filedialog
# import os
#
# from win32com.client import Dispatch
#
# def open_file():
#     # Create and hide the root window
#     root = tk.Tk()
#     root.withdraw()
#
#     # Open file dialog and get the selected file path
#     file_path = filedialog.askopenfilename(
#         title="Select File",
#         filetypes=[("All Files", "*.*"),
#                    ("Word Documents", "*.doc;*.docx"),
#                    ("Excel Files", "*.xls;*.xlsx")]
#     )
#
#     if file_path:
#         # Open Word documents
#         if file_path.endswith(('.doc', '.docx')):
#             word = Dispatch("Word.Application")
#             word.Visible = True
#             word.Documents.Open(file_path)
#         # Open Excel spreadsheets
#         elif file_path.endswith(('.xls', '.xlsx')):
#             excel = Dispatch("Excel.Application")
#             excel.Visible = True
#             excel.Workbooks.Open(file_path)
#         # Open with the default application for other files
#         else:
#             os.startfile(file_path)
#
# if __name__ == "__main__":
#     open_file()
import tkinter as tk
from tkinter import filedialog
import os


def open_file():
    # Create and hide the root window
    root = tk.Tk()
    root.withdraw()

    # Open file dialog and get the selected file path
    file_path = filedialog.askopenfilename(
        title="Select File",
        filetypes=[("All Files", "*.*"),
                   ("Word Documents", "*.doc;*.docx"),
                   ("Excel Files", "*.xls;*.xlsx")]
    )

    if file_path:
        # Open with the default application for all files
        os.startfile(file_path)


if __name__ == "__main__":
    open_file()
