from pyzbar.pyzbar import decode
import fitz
import pandas as pd
import os
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class PDFProcessorApp:
    def __init__(self, master):
        self.master = master
        self.master.title("PDF Processor App")

        self.pdf_path = None

        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.image_path = os.path.join(script_dir, "bp logo.png")

        original_image = Image.open(self.image_path)
        resized_image = original_image.resize((180, 120), Image.ANTIALIAS if hasattr(Image, 'ANTIALIAS') else Image.BICUBIC)
        self.image = ImageTk.PhotoImage(resized_image)

        self.label = tk.Label(self.master, image=self.image)
        self.label.grid(row=0, column=0, pady=30, padx=(150,0), sticky='nsew')

        self.file_label = tk.Label(self.master, text="Nenhum PDF selecionado.", wraplength=200)
        self.file_label.grid(row=1, column=0, pady=10, padx=(150, 0), sticky='nsew')

        self.select_pdf_button = ttk.Button(self.master, text="Selecionar PDF", command=self.select_pdf)
        self.select_pdf_button.grid(row=2, column=0, pady=7, padx=(150,0), sticky='nsew')

        self.process_pdf_button = ttk.Button(self.master, text="Processar PDF", command=self.process_pdf)
        self.process_pdf_button.grid(row=3, column=0, pady=7, padx=(150,0), sticky='nsew')

        self.center_window()

    def center_window(self):
        w = 500
        h = 400
        ws = self.master.winfo_screenwidth()
        hs = self.master.winfo_screenheight()
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        self.master.geometry('%dx%d+%d+%d' % (w, h, x, y))

    def select_pdf(self):
        self.pdf_path = filedialog.askopenfilename(title="Selecione o arquivo PDF", filetypes=[("PDF Files", "*.pdf")])
        if not self.pdf_path:
            self.file_label.config(text="Nenhum PDF selecionado.")
        else:
            self.file_label.config(text=f"PDF selecionado: {os.path.basename(self.pdf_path)}")

    def extract_qr_codes_from_page(self, page, page_number):
        transform = fitz.Matrix(1, 1).prescale(2, 2)
        pixmap = page.get_pixmap(matrix=transform)
        image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
        decoded_objects = decode(image)
        qr_codes_data = []

        for obj in decoded_objects:
            try:
                qr_data = obj.data.decode('utf-8')
            except Exception as e:
                qr_data = f'Erro na decodificação: {str(e)} (Página {page_number}, Posição X: {obj.rect.left}, Posição Y: {obj.rect.top})'
            qr_codes_data.append(qr_data)

        if not qr_codes_data:
            qr_codes_data.append(f'Nenhum QR Code encontrado (Pág. {page_number+1})')

        return qr_codes_data

    def process_pdf(self):
        if not self.pdf_path:
            self.file_label.config(text="Nenhum arquivo PDF selecionado.")
            return

        pdf_document = fitz.open(self.pdf_path)
        qr_codes_data_list = []

        for page_number in range(pdf_document.page_count):
            page = pdf_document[page_number]
            qr_codes_data = self.extract_qr_codes_from_page(page, page_number)
            qr_codes_data_list.extend(qr_codes_data)

        df = pd.DataFrame({"QR Code Data": qr_codes_data_list})

        pdf_folder, pdf_filename = os.path.split(self.pdf_path)
        output_excel_path = os.path.join(pdf_folder, 'informacoes_qr_codes.xlsx')
        df.to_excel(output_excel_path, index=False)

        messagebox.showinfo("Sucesso", f'Dados dos QR Codes salvos com sucesso. Excel salvo em {output_excel_path}')

        self.pdf_path = None
        self.file_label.config(text="Nenhum arquivo PDF selecionado.")
        self.label.config(text="")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.geometry("500x400")
    root.mainloop()
