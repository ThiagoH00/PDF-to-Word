import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
from pdf2docx import Converter

def run_conversion_in_background(app_instance, pdf_path):
    """
    Função que executa a conversão e atualiza a GUI com o resultado.
    É executada em uma thread separada para não travar a interface.
    """
    try:
        # Define o nome do arquivo de saída (mesmo nome, extensão .docx)
        output_path, _ = os.path.splitext(pdf_path)
        output_path += ".docx"

        # Cria o objeto conversor e faz a conversão
        cv = Converter(pdf_path)
        cv.convert(output_path, start=0, end=None)  # type: ignore
        cv.close()

        # Atualiza a GUI na thread principal com mensagem de sucesso
        app_instance.root.after(0, app_instance.on_conversion_success, output_path)

    except Exception as e:
        # Atualiza a GUI na thread principal com mensagem de erro
        app_instance.root.after(0, app_instance.on_conversion_error, e)

class PdfConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor PDF para Word")
        self.root.geometry("500x230")
        self.root.resizable(False, False)

        self.pdf_path = ""

        # --- Widgets ---
        self.frame = tk.Frame(root, padx=15, pady=15)
        self.frame.pack(expand=True, fill=tk.BOTH)

        self.label_info = tk.Label(self.frame, text="Selecione o arquivo PDF para converter:", justify=tk.LEFT)
        self.label_info.pack(pady=5, anchor="w")

        self.btn_select = tk.Button(self.frame, text="Selecionar Arquivo PDF", command=self.select_file)
        self.btn_select.pack(pady=5, fill=tk.X)

        self.label_file = tk.Label(self.frame, text="Nenhum arquivo selecionado.", fg="gray", wraplength=480)
        self.label_file.pack(pady=10, anchor="w")

        self.btn_convert = tk.Button(self.frame, text="Converter para Word", command=self.start_conversion, state=tk.DISABLED)
        self.btn_convert.pack(pady=10, ipady=4, fill=tk.X)

    def select_file(self):
        """Abre a caixa de diálogo para selecionar um arquivo PDF."""
        path = filedialog.askopenfilename(
            title="Selecione um arquivo PDF",
            filetypes=(("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*"))
        )
        if path:
            self.pdf_path = path
            filename = os.path.basename(path)
            self.label_file.config(text=f"Arquivo: {filename}", fg="black")
            self.btn_convert.config(state=tk.NORMAL)

    def start_conversion(self):
        """Inicia o processo de conversão em uma thread separada."""
        if not self.pdf_path:
            messagebox.showwarning("Aviso", "Por favor, selecione um arquivo PDF primeiro.")
            return

        self.btn_convert.config(state=tk.DISABLED, text="Convertendo...")
        self.btn_select.config(state=tk.DISABLED)

        # Executa a conversão em uma nova thread para não travar a GUI
        thread = threading.Thread(target=run_conversion_in_background, args=(self, self.pdf_path))
        thread.start()

    def on_conversion_success(self, output_path):
        messagebox.showinfo("Sucesso", f"Arquivo convertido com sucesso!\nSalvo como: {output_path}")
        self.btn_convert.config(state=tk.NORMAL, text="Converter para Word")
        self.btn_select.config(state=tk.NORMAL)

    def on_conversion_error(self, error):
        messagebox.showerror("Erro de Conversão", f"Ocorreu um erro:\n{error}")
        self.btn_convert.config(state=tk.NORMAL, text="Converter para Word")
        self.btn_select.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = PdfConverterApp(root)
    root.mainloop()
