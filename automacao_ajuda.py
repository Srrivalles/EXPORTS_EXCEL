import os
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, ttk, messagebox
from pandastable import Table
import time
from fpdf import FPDF

class ExcelManager(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Excel Data Manager")
        self.geometry("1200x700")
        self.minsize(1000, 600)
        
        ctk.set_appearance_mode("dark")  # Default Theme
        ctk.set_default_color_theme("blue")
        
        self.file_path = None
        self.dataframe = None
        self.recent_files = []
        
        self.create_loading_screen()
    
    def create_loading_screen(self):
        self.loading_frame = ctk.CTkFrame(self)
        self.loading_frame.pack(fill="both", expand=True)
        
        self.loading_label = ctk.CTkLabel(self.loading_frame, text="Seja Bem-Vindo!", font=("Arial", 24))
        self.loading_label.pack(pady=20)
        
        self.progress = ctk.CTkProgressBar(self.loading_frame)
        self.progress.pack(pady=10)
        self.progress.set(0)
        
        self.update_idletasks()
        for i in range(100):
            time.sleep(0.03)
            self.progress.set(i / 100)
            self.update_idletasks()
        
        self.loading_frame.destroy()
        self.create_widgets()
    
    def create_widgets(self):
        # Frame superior
        self.top_frame = ctk.CTkFrame(self)
        self.top_frame.pack(fill="x", padx=10, pady=5)
        
        self.open_button = ctk.CTkButton(self.top_frame, text="📂 Abrir Arquivo", command=self.load_file)
        self.open_button.pack(side="left", padx=5)
        
        self.theme_switch = ctk.CTkSwitch(self.top_frame, text="Modo Escuro", command=self.toggle_theme)
        self.theme_switch.pack(side="left", padx=5)
        
        self.export_button = ctk.CTkButton(self.top_frame, text="📄 Exportar PDF", command=self.export_to_pdf)
        self.export_button.pack(side="left", padx=5)
        
        # Frame da Tabela
        self.table_frame = ctk.CTkFrame(self)
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.table = None
    
    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not file_path:
            return
        
        self.file_path = file_path
        self.dataframe = pd.read_excel(file_path)
        
        # Remove colunas sem título
        self.dataframe = self.dataframe.loc[:, ~self.dataframe.columns.str.contains("^Unnamed")]
        
        self.update_table()
    
    def update_table(self):
        if self.table:
            self.table.destroy()
        
        self.table = Table(self.table_frame, dataframe=self.dataframe, showtoolbar=True, showstatusbar=True)
        self.table.show()
    
    def toggle_theme(self):
        new_mode = "dark" if ctk.get_appearance_mode() == "light" else "light"
        ctk.set_appearance_mode(new_mode)
    
    def export_to_pdf(self):
        if self.dataframe is None:
            messagebox.showerror("Erro", "Nenhuma planilha carregada!")
            return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not file_path:
            return
        
        try:
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.cell(200, 10, txt="Relatório de Dados", ln=True, align='C')
            
            for col in self.dataframe.columns:
                pdf.cell(40, 10, col, border=1)
            pdf.ln()
            
            for i in range(min(30, len(self.dataframe))):  # Limitando para evitar excesso de páginas
                for col in self.dataframe.columns:
                    pdf.cell(40, 10, str(self.dataframe.iloc[i][col]), border=1)
                pdf.ln()
            
            pdf.output(file_path)
            messagebox.showinfo("Sucesso", "Relatório exportado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exportar PDF: {str(e)}")

if __name__ == "__main__":
    app = ExcelManager()
    app.mainloop()
