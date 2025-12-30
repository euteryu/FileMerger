import os
import threading
import datetime
import win32com.client
from pypdf import PdfWriter

# Valid extensions configuration
VALID_EXTS = {'.pptx', '.ppt', '.docx', '.doc', '.pdf'}

class MergeWorker(threading.Thread):
    def __init__(self, file_data_list, output_path, progress_callback, status_callback, done_callback):
        super().__init__()
        self.file_data_list = file_data_list 
        self.output_path = output_path
        self.progress_callback = progress_callback
        self.status_callback = status_callback
        self.done_callback = done_callback

    def run(self):
        try:
            self.merge_to_pdf()
            self.done_callback(True, "PDF Created Successfully!")
        except Exception as e:
            self.done_callback(False, str(e))

    def get_temp_pdf_name(self, original_path, index):
        folder = os.path.dirname(original_path)
        return os.path.join(folder, f"~temp_merger_{index}.pdf")

    def merge_to_pdf(self):
        merger = PdfWriter()
        temp_files = []
        
        ppt_app = None
        word_app = None
        
        try: ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        except: pass

        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
        except: pass

        total = len(self.file_data_list)

        for i, item in enumerate(self.file_data_list):
            filepath = item['path']
            filename = os.path.basename(filepath)
            ext = os.path.splitext(filepath)[1].lower()
            abs_path = os.path.abspath(filepath)
            
            self.progress_callback((i / total))
            self.status_callback(f"Processing ({i+1}/{total}): {filename}")

            if ext == '.pdf':
                try:
                    merger.append(abs_path)
                except Exception as e:
                    print(f"Error appending PDF {filename}: {e}")
                continue

            temp_pdf = self.get_temp_pdf_name(abs_path, i)
            converted = False

            try:
                if ext in ['.ppt', '.pptx']:
                    if not ppt_app: raise Exception("PowerPoint not found.")
                    deck = ppt_app.Presentations.Open(abs_path, ReadOnly=True, WithWindow=False)
                    deck.SaveAs(temp_pdf, 32) # ppSaveAsPDF
                    deck.Close()
                    converted = True
                
                elif ext in ['.doc', '.docx']:
                    if not word_app: raise Exception("Word not found.")
                    doc = word_app.Documents.Open(abs_path, ReadOnly=True)
                    doc.ExportAsFixedFormat(temp_pdf, 17) # wdExportFormatPDF
                    doc.Close(False)
                    converted = True

                if converted and os.path.exists(temp_pdf):
                    merger.append(temp_pdf)
                    temp_files.append(temp_pdf)

            except Exception as e:
                print(f"Failed to convert {filename}: {e}")

        self.status_callback("Finalizing PDF file...")
        merger.write(self.output_path)
        merger.close()

        for f in temp_files:
            try: os.remove(f)
            except: pass
        
        if ppt_app: 
            try: ppt_app.Quit()
            except: pass
        if word_app: 
            try: word_app.Quit() 
            except: pass

def get_file_metadata(filepath):
    try:
        timestamp = os.path.getmtime(filepath)
        dt = datetime.datetime.fromtimestamp(timestamp)
        date_str = dt.strftime("%Y-%m-%d %H:%M")
        ext = os.path.splitext(filepath)[1].lower()
        type_map = {'.docx':'WORD', '.doc':'WORD', '.pptx':'PPT', '.ppt':'PPT', '.pdf':'PDF'}
        file_type = type_map.get(ext, "FILE")
        return {"date": date_str, "timestamp": timestamp, "type": file_type, "name": os.path.basename(filepath)}
    except:
        return {"date": "-", "timestamp": 0, "type": "FILE", "name": os.path.basename(filepath)}

def is_valid_file(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    return ext in VALID_EXTS and not os.path.basename(filepath).startswith('~$')