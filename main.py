import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import os
from datetime import datetime

class InventoryControlSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Controle de Estoque de Máquinas")
        self.root.geometry("800x600")
        
        self.mercado_pago_file = ""
        self.hunting_instore_file = ""
        
        self.consolidated_data = None
        
        self.create_widgets()
    
    def create_widgets(self):
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        title_label = tk.Label(main_frame, text="Controle de Estoque de Máquinas", 
                             font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        upload_frame = tk.LabelFrame(main_frame, text="Importar Arquivos", padx=10, pady=10)
        upload_frame.pack(fill=tk.X, pady=10)
        
        mercado_pago_btn = tk.Button(upload_frame, text="1. Selecionar Arquivo Mercado Pago (Excel)", 
                                    command=self.load_mercado_pago)
        mercado_pago_btn.pack(fill=tk.X, pady=5)
        self.mercado_pago_label = tk.Label(upload_frame, text="Nenhum arquivo selecionado", fg="gray")
        self.mercado_pago_label.pack(fill=tk.X)
        
        hunting_instore_btn = tk.Button(upload_frame, text="2. Selecionar Arquivo Hunting Instore (CSV/Excel)", 
                                       command=self.load_hunting_instore)
        hunting_instore_btn.pack(fill=tk.X, pady=5)
        self.hunting_instore_label = tk.Label(upload_frame, text="Nenhum arquivo selecionado", fg="gray")
        self.hunting_instore_label.pack(fill=tk.X)
        
        process_btn = tk.Button(main_frame, text="Processar Dados", 
                               command=self.process_data, 
                               bg="#4CAF50", fg="white", font=("Arial", 12))
        process_btn.pack(pady=20)
        
        self.status_label = tk.Label(main_frame, text="Aguardando arquivos...", fg="blue")
        self.status_label.pack(fill=tk.X)
        
        self.export_btn = tk.Button(main_frame, text="Exportar Dados Consolidados", 
                                   command=self.export_data, 
                                   state=tk.DISABLED, bg="#2196F3", fg="white")
        self.export_btn.pack(pady=10)
        
        self.create_data_preview(main_frame)
    
    def create_data_preview(self, parent):
        preview_frame = tk.LabelFrame(parent, text="Pré-visualização dos Dados", padx=10, pady=10)
        preview_frame.pack(expand=True, fill=tk.BOTH)
        
        y_scroll = tk.Scrollbar(preview_frame)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        x_scroll = tk.Scrollbar(preview_frame, orient=tk.HORIZONTAL)
        x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.tree = ttk.Treeview(preview_frame, 
                                 yscrollcommand=y_scroll.set, 
                                 xscrollcommand=x_scroll.set)
        self.tree.pack(expand=True, fill=tk.BOTH)
        
        y_scroll.config(command=self.tree.yview)
        x_scroll.config(command=self.tree.xview)
    
    def load_mercado_pago(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo Mercado Pago",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.mercado_pago_file = file_path
            self.mercado_pago_label.config(text=os.path.basename(file_path), fg="green")
            self.update_status()
    
    def load_hunting_instore(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo Hunting Instore",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.hunting_instore_file = file_path
            self.hunting_instore_label.config(text=os.path.basename(file_path), fg="green")
            self.update_status()
    
    def update_status(self):
        if self.mercado_pago_file and self.hunting_instore_file:
            self.status_label.config(text="Arquivos prontos para processamento", fg="green")
        elif self.mercado_pago_file or self.hunting_instore_file:
            self.status_label.config(text="Falta selecionar um arquivo", fg="orange")
        else:
            self.status_label.config(text="Aguardando arquivos...", fg="blue")
    
    def process_data(self):
        if not self.mercado_pago_file or not self.hunting_instore_file:
            messagebox.showerror("Erro", "Por favor, selecione ambos os arquivos antes de processar.")
            return
        
        try:
            mercado_pago_data = self.read_mercado_pago()
            hunting_instore_data = self.read_hunting_instore()
            self.consolidate_data(mercado_pago_data, hunting_instore_data)
            self.update_data_preview()
            self.export_btn.config(state=tk.NORMAL)
            messagebox.showinfo("Sucesso", "Dados processados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao processar os dados:\n{str(e)}")
    
    def read_mercado_pago(self):
        xls = pd.ExcelFile(self.mercado_pago_file)
        
        avanco_df = pd.read_excel(xls, sheet_name="AVANÇO")
        avanco_df = avanco_df.rename(columns={'DATA': 'data_envio', 'SERIAL': 'serial', 'SUPERVISOR': 'supervisor'})
        avanco_df['status'] = 'ENVIADA'
        avanco_df['origem'] = 'AVANÇO'
        
        quebradas_df = pd.read_excel(xls, sheet_name="QUEBRADAS")
        quebradas_df = quebradas_df.rename(columns={'DATA': 'data_quebra', 'SERIAL': 'serial', 'MODELO': 'modelo_defeito'})
        quebradas_df['status'] = 'QUEBRADA'
        quebradas_df['origem'] = 'QUEBRADAS'
        
        entrada_df = pd.read_excel(xls, sheet_name="ENTRADA")
        entrada_df = entrada_df.rename(columns={'DATA': 'data_entrada', 'SERIAL': 'serial', 'ID': 'caixa'})
        entrada_df['status'] = 'ESTOQUE'
        entrada_df['origem'] = 'ENTRADA'
        
        mercado_pago_consolidado = pd.concat([
            avanco_df[['serial', 'data_envio', 'supervisor', 'status', 'origem']],
            quebradas_df[['serial', 'data_quebra', 'modelo_defeito', 'status', 'origem']],
            entrada_df[['serial', 'data_entrada', 'caixa', 'status', 'origem']]
        ], ignore_index=True)
        
        return mercado_pago_consolidado
    
    def read_hunting_instore(self):
        ext = os.path.splitext(self.hunting_instore_file)[1].lower()
        
        if ext == '.csv':
            hunting_df = pd.read_csv(self.hunting_instore_file, encoding='utf-8', sep=',')
        else:
            hunting_df = pd.read_excel(self.hunting_instore_file)
        
        hunting_df = hunting_df.rename(columns={
            'SN Device': 'serial',
            'Data Venda': 'data_venda',
            'Modelo Device': 'modelo_vendido'
        })
        
        hunting_df['status'] = 'VENDIDA'
        hunting_df['origem'] = 'HUNTING_INSTORE'
        
        return hunting_df[['serial', 'data_venda', 'modelo_vendido', 'status', 'origem']]
    
    def consolidate_data(self, mercado_pago_data, hunting_instore_data):
        # Padronizar os seriais para os últimos 12 caracteres e em maiúsculas
        mercado_pago_data['serial'] = mercado_pago_data['serial'].astype(str).str.upper().str[-12:]
        hunting_instore_data['serial'] = hunting_instore_data['serial'].astype(str).str.upper().str[-12:]

        # Juntar os dois datasets
        all_data = pd.concat([mercado_pago_data, hunting_instore_data], ignore_index=True)

        # Mapeamento de prioridade
        status_priority = {
            "VENDIDA": 1,
            "QUEBRADA": 2,
            "ENVIADA": 3,
            "ESTOQUE": 4
        }

        all_data["prioridade_status"] = all_data["status"].map(status_priority)
        all_data = all_data.sort_values(by=["serial", "prioridade_status"])
        deduplicado = all_data.groupby("serial", as_index=False).first()
        deduplicado.drop(columns=["prioridade_status"], inplace=True)

        self.consolidated_data = deduplicado
        self.consolidated_data = self.consolidated_data.sort_values(by=["status", "serial"], ascending=[True, True])
    
    def update_data_preview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if self.consolidated_data is None or self.consolidated_data.empty:
            return
        
        preview_df = self.consolidated_data.head(100)
        
        self.tree["columns"] = list(preview_df.columns)
        self.tree["show"] = "headings"
        
        for col in preview_df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.CENTER)
        
        for _, row in preview_df.iterrows():
            self.tree.insert("", tk.END, values=list(row))
    
    def export_data(self):
        if self.consolidated_data is None or self.consolidated_data.empty:
            messagebox.showerror("Erro", "Nenhum dado para exportar. Processe os dados primeiro.")
            return
        
        default_filename = f"consolidado_estoque_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        file_path = filedialog.asksaveasfilename(
            title="Salvar arquivo consolidado",
            initialfile=default_filename,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.consolidated_data.to_excel(file_path, index=False)
                messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao salvar arquivo:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryControlSystem(root)
    root.mainloop()
