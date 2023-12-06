import sys
import os
from pdf2image import convert_from_path
import fitz
import easyocr
from docx import Document
from docx.shared import Pt
import wx
from threading import Thread
from io import BytesIO

# Get the current working directory
current_directory = os.getcwd()

# Add the poppler bin directory to the system PATH
poppler_bin_path = os.path.join(current_directory, 'Library', 'bin')
os.environ['PATH'] += ';' + poppler_bin_path


class ConverterThread(Thread):
    def __init__(self, pdf_path, docx_path):
        super().__init__()
        self.pdf_path = pdf_path
        self.docx_path = docx_path

    def run(self):
        self.convert_pdf_to_docx()
    
    def calculate_average_x_position(left_x, right_x):
        # Calculate the average x position of the text
        return (left_x + right_x) / 2
    
    def convert_pdf_to_docx(self):
        try:
            document = Document()
            results_list = []
            # Initialize EasyOCR reader
            reader = easyocr.Reader(['ru','en'])
            images = convert_from_path(self.pdf_path)
            
            for page_number, image in enumerate(images):
                # Convert PIL Image to bytes
                img_byte_array = BytesIO()
                image.save(img_byte_array, format='PNG')
                img_bytes = img_byte_array.getvalue()

                # Use EasyOCR to read text from the image
                results = reader.readtext(img_bytes)
                results_list.append(results)
                                
            for page_number, results in enumerate(results_list):
                for result in results:
                    bbox, text, prob = result
                    document.add_paragraph(f'{text}'   )
            document.save(self.docx_path)
        except Exception as e:
            print(f"Error during conversion: {e}")

class PDFConverterApp(wx.Frame):
    def __init__(self):
        super().__init__(None, title='PDF to DOCX Converter', size=(400, 120))

        self.converter_thread = None

        self.init_ui()

    def init_ui(self):
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.status_label = wx.StaticText(panel, label='Status: Ready')
        sizer.Add(self.status_label, 0, wx.ALL | wx.CENTER, 10)

        self.convert_button = wx.Button(panel, label='Convert PDF to DOCX')
        self.convert_button.Bind(wx.EVT_BUTTON, self.convert_pdf_to_docx)
        sizer.Add(self.convert_button, 0, wx.ALL | wx.CENTER, 10)

        panel.SetSizer(sizer)

    def convert_pdf_to_docx(self, event):
        with wx.FileDialog(self, "Select PDF File", wildcard="PDF Files (*.pdf)|*.pdf", style=wx.FD_OPEN) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_OK:
                pdf_path = file_dialog.GetPath()

                with wx.FileDialog(self, "Save As", wildcard="DOCX Files (*.docx)|*.docx", style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as save_dialog:
                    if save_dialog.ShowModal() == wx.ID_OK:
                        docx_path = save_dialog.GetPath()
                        self.converter_thread = ConverterThread(pdf_path, docx_path)
                        self.status_label.SetLabel('Status: Converting...')
                        self.converter_thread.start()
                        self.converter_thread.join()  # Wait for the thread to finish

                        self.status_label.SetLabel('Status: Conversion Finished')

if __name__ == '__main__':
    app = wx.App(False)
    frame = PDFConverterApp()
    frame.Show()
    app.MainLoop()
