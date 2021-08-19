# pip install pywin32

import os
import win32com.client as win32


def merge_docx(source_files_path=r"D:\homework\all_docxs",
               result_file=r"D:\homework\new_folder\result.docx"):
    """
    :param source_files_path: all docxs file path
    :param result_file: target docxs file path
    :return:
    """
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    output = word.Documents.Add()
    source_docx_path = os.listdir(source_files_path)
    source_docxs = []
    for file in source_docx_path:
        if file.endswith(".docx") or file.endswith(".doc"):
            source_docxs.append(source_files_path + "\\" + file)
    i = 0
    source_docxs.reverse()
    for source_docx in source_docxs:
        source_docx_filename_list = (source_docx.split("\\"))[-1].split(".")
        print(source_docx_filename_list[-2])
        if i >= 1:
            output.Application.Selection.Range.InsertBreak()
        output.Application.Selection.Range.InsertFile(source_docx)
        i += 1
    output.SaveAs(result_file)
    output.Close()

if __name__ == '__main__':
    merge_docx()
