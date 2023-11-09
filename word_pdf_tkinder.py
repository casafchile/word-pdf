import tkinter as tk
from tkinter import filedialog
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def convert_docx_to_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    
    if file_path:
        pdf_file_path = file_path.replace(".docx", ".pdf")
        doc = SimpleDocTemplate(pdf_file_path, pagesize=letter)
        styles = getSampleStyleSheet()
        pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
        
        try:
            docx = Document(file_path)
            pdf_content = []
            
            for para in docx.paragraphs:
                text = para.text
                p = Paragraph(text, styles['Normal'])
                pdf_content.append(p)
            
            doc.build(pdf_content)
            result_label.config(text="Conversión exitosa: " + pdf_file_path)
        except Exception as e:
            result_label.config(text="Error en la conversión: " + str(e))

app = tk.Tk()

app.title("Convertidor de DOCX a PDF")


convert_button = tk.Button(app, text="Convertir DOCX a PDF", command=convert_docx_to_pdf)
convert_button.pack(pady=20)

result_label = tk.Label(app, text="")
result_label.pack()

app.mainloop()
