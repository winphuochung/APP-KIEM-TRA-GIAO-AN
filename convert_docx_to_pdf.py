import os
try:
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc_path = r"D:\APP-KIEM-TRA-GIAO-AN\BAO_CAO_CHI_TIET_TIEN_DO_9_CHU_KY_TUAN_GOOGLE_DRIVE.docx"
    pdf_path = r"D:\APP-KIEM-TRA-GIAO-AN\BAO_CAO_CHI_TIET_TIEN_DO_9_CHU_KY_TUAN_GOOGLE_DRIVE.pdf"
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    word.Quit()
    print("PDF converted successfully via win32com!")
except Exception as e:
    print("PDF conversion error:", e)
