import os
from pdf2image import convert_from_path
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

    def convert_pdf_to_docx(self):
        try:
            document = Document()
            results_list = []
            # Initialize EasyOCR reader
            reader = easyocr.Reader(['ru', 'en'])
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
                    # Create a paragraph with position information
                    paragraph = document.add_paragraph(f'{text}')

                    # Get the last run in the paragraph (assuming there is one)
                    current_run = paragraph.runs[-1]

                    # Set the font size
                    current_run.font.size = Pt(12)  # You can adjust the font size as needed

                    # Set the position of the paragraph on the page
                    paragraph._element.pPr.get_or_add_position().val = str(bbox[0][0])  # X-coordinate

                    # Save the document
                    document.save(self.docx_path)

        except Exception as e:
            print(f"Error during conversion: {e}")

class PDFConverterApp:
    def __init__(self):
        self.app = wx.App(False)
        self.frame = None
        self.converter_thread = None
        self.init_ui()

    def init_ui(self):
        self.frame = wx.Frame(None, title='PDF to DOCX Converter', size=(400, 120))
        panel = wx.Panel(self.frame)
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.status_label = wx.StaticText(panel, label='Status: Ready')
        sizer.Add(self.status_label, 0, wx.ALL | wx.CENTER, 10)

        self.convert_button = wx.Button(panel, label='Convert PDF to DOCX')
        self.convert_button.Bind(wx.EVT_BUTTON, self.convert_pdf_to_docx)
        sizer.Add(self.convert_button, 0, wx.ALL | wx.CENTER, 10)

        panel.SetSizer(sizer)

    def convert_pdf_to_docx(self, event):
        with wx.FileDialog(None, "Select PDF File", wildcard="PDF Files (*.pdf)|*.pdf", style=wx.FD_OPEN) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_OK:
                pdf_path = file_dialog.GetPath()

                with wx.FileDialog(None, "Save As", wildcard="DOCX Files (*.docx)|*.docx",
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as save_dialog:
                    if save_dialog.ShowModal() == wx.ID_OK:
                        docx_path = save_dialog.GetPath()
                        self.converter_thread = ConverterThread(pdf_path, docx_path)
                        self.status_label.SetLabel('Status: Converting...')
                        self.converter_thread.start()
                        self.converter_thread.join()  # Wait for the thread to finish
                        self.status_label.SetLabel('Status: Conversion Finished')

    def run(self):
        self.frame.Show()
        self.app.MainLoop()

if __name__ == '__main__':
    app = PDFConverterApp()
    app.run()
