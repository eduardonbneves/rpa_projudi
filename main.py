import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import json
import os
import threading
import datetime
import pandas as pd
import automation

CONFIG_FILE = "config.json"

class RPAApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Habilitação PROJUDI AM")
        self.root.geometry("500x750")
        self.root.resizable(True, True)

        # Style configuration
        self.style = ttk.Style()
        self.style.theme_use('alt')
        
        # Colors and Fonts
        bg_color = "#f0f0f0"
        self.root.configure(bg=bg_color)
        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabel", background=bg_color, font=("Segoe UI", 10))
        self.style.configure("TButton", font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"))
        self.style.configure("Status.TLabel", font=("Segoe UI", 10))
        self.style.configure("Green.Horizontal.TProgressbar",
            background="#4CAF50",       # Cor da barra (Verde)
            troughcolor="#E0E0E0",      # Cor do fundo (Cinza claro)
            bordercolor="#E0E0E0",      # Borda para ficar 'flat'
            lightcolor="#4CAF50",       # Sombras (mesma cor para ficar flat)
            darkcolor="#4CAF50")        # Sombras

        # Variables
        self.usuario_var = tk.StringVar()
        self.senha_var = tk.StringVar()
        self.oab_var = tk.StringVar()
        self.excel_path_var = tk.StringVar()
        self.report_path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Pronto para executar")
        self.progress_var = tk.DoubleVar(value=0)
        
        self.stop_flag = False
        self.is_running = False
        self.current_log_file = None
        self.sucessos = []

        self.create_widgets()
        self.load_config()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Configurações de Acesso ---
        ttk.Label(main_frame, text="🔐 Configurações de Acesso", style="Header.TLabel").pack(anchor="w", pady=(0, 10))
        
        access_frame = ttk.Frame(main_frame)
        access_frame.pack(fill=tk.X, pady=(0, 15))

        self._create_entry(access_frame, "Usuário PROJUDI:", self.usuario_var)
        self._create_password_entry(access_frame, "Senha:", self.senha_var)
        self._create_entry(access_frame, "Número da OAB:", self.oab_var)

        save_btn = ttk.Button(access_frame, text="💾 Salvar Configurações", command=self.save_config)
        save_btn.pack(pady=(10, 0))
        
        ttk.Label(access_frame, text="ℹ As configurações serão salvas automaticamente", font=("Segoe UI", 8), foreground="#666").pack(pady=(5, 0))

        # --- Arquivo de Processos ---
        ttk.Label(main_frame, text="📄 Arquivo de Processos", style="Header.TLabel").pack(anchor="w", pady=(10, 10))
        
        process_frame = ttk.Frame(main_frame)
        process_frame.pack(fill=tk.X, pady=(0, 15))

        self._create_file_selector(process_frame, "Arquivo Excel:", self.excel_path_var, self.select_excel)
        
        ttk.Label(process_frame, text="ℹ O arquivo deve ter uma coluna chamada 'PROCESSOS'", font=("Segoe UI", 8), foreground="#666").pack(anchor="w", pady=(5, 0))

        # --- Local do Relatório ---
        ttk.Label(main_frame, text="📂 Local do Relatório", style="Header.TLabel").pack(anchor="w", pady=(10, 5))
        report_frame = ttk.Frame(main_frame)
        report_frame.pack(fill=tk.X, pady=(0, 10))
        self._create_file_selector(report_frame, "Pasta Destino:", self.report_path_var, self.select_folder)
        
        ttk.Label(report_frame, text="ℹ Se não selecionado, será salvo na pasta Downloads", font=("Segoe UI", 8), foreground="#666").pack(anchor="w", pady=(5, 0))

        # --- Execução ---
        ttk.Label(main_frame, text="▶ Execução", style="Header.TLabel").pack(anchor="w", pady=(10, 10))
        
        exec_frame = ttk.Frame(main_frame)
        exec_frame.pack(fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(exec_frame)
        btn_frame.pack(pady=(0, 10))

        self.start_btn = ttk.Button(btn_frame, text="🚀 Iniciar Habilitação", command=self.start_automation)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(btn_frame, text="⏹ Parar", command=self.stop_automation, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        # Status and Progress
        self.status_label = ttk.Label(exec_frame, textvariable=self.status_var, style="Status.TLabel", anchor="center")
        self.status_label.pack(fill=tk.X, pady=(10, 5))

        self.progress_bar = ttk.Progressbar(exec_frame, variable=self.progress_var, maximum=100, style="Green.Horizontal.TProgressbar")
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))

        # Log Viewer
        ttk.Label(exec_frame, text="📜 Logs de Execução", style="Header.TLabel").pack(anchor="w", pady=(10, 5))
        self.log_text = ScrolledText(exec_frame, height=10, state='disabled', font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _create_entry(self, parent, label_text, variable):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        ttk.Label(frame, text=label_text, width=20).pack(side=tk.LEFT)
        ttk.Entry(frame, textvariable=variable).pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _create_password_entry(self, parent, label_text, variable):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        ttk.Label(frame, text=label_text, width=20).pack(side=tk.LEFT)
        
        self.pass_entry = ttk.Entry(frame, textvariable=variable, show="*")
        self.pass_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Toggle visibility button (simple text for now, could be icon)
        self.show_pass = False
        self.toggle_btn = ttk.Button(frame, text="👁", width=3, command=self.toggle_password)
        self.toggle_btn.pack(side=tk.LEFT, padx=(5, 0))

    def _create_file_selector(self, parent, label_text, variable, command):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        ttk.Label(frame, text=label_text, width=15).pack(side=tk.LEFT)
        ttk.Entry(frame, textvariable=variable, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(frame, text="📂 Selecionar", command=command).pack(side=tk.LEFT, padx=(5, 0))

    def toggle_password(self):
        self.show_pass = not self.show_pass
        self.pass_entry.configure(show="" if self.show_pass else "*")

    def select_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            self.excel_path_var.set(filename)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.report_path_var.set(folder)

    def save_config(self):
        config = {
            "usuario": self.usuario_var.get(),
            "senha": self.senha_var.get(),
            "oab": self.oab_var.get(),
            "excel_path": self.excel_path_var.get(),
            "report_path": self.report_path_var.get()
        }
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(config, f)
            messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar configurações: {e}")

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    config = json.load(f)
                    self.usuario_var.set(config.get("usuario", ""))
                    self.senha_var.set(config.get("senha", ""))
                    self.oab_var.set(config.get("oab", ""))
                    self.excel_path_var.set(config.get("excel_path", ""))
                    self.report_path_var.set(config.get("report_path", ""))
            except Exception:
                pass # Ignore errors loading config

    def start_automation(self):
        if self.is_running: return
        excel = self.excel_path_var.get()
        
        # Define onde salvar o relatório
        folder = self.report_path_var.get()
        if not folder or not os.path.exists(folder):
           folder = os.path.join(os.path.expanduser("~"), "Downloads")
            
        if not excel: return messagebox.showwarning("Aviso", "Selecione o Excel.")
        
        procs, erro = automation.ler_processos_excel(excel)
        if erro: return messagebox.showerror("Erro", erro)
        if not procs: return messagebox.showwarning("Aviso", "Lista vazia.")

        # --- PREPARAÇÃO DO ARQUIVO DE LOG ---
        timestamp_nome = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.current_log_file = os.path.join(folder, f"logs_Projudi_{timestamp_nome}.txt")
        
        self.sucessos = []

        # Cria o arquivo e coloca cabeçalho
        with open(self.current_log_file, "w", encoding="utf-8") as f:
            f.write(f"LOGS DE EXECUÇÃO - PROJUDI\n")
            f.write(f"Data: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}\n")
            f.write(f"Total de processos: {len(procs)}\n")
            f.write("-" * 50 + "\n")
        # ------------------------------------

        self.is_running = True
        self.stop_flag = False
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.progress_var.set(0)
        
        # Limpa visualizador da tela
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        self.log_message(f"Arquivo de logs criado em: {self.current_log_file}")

        thread = threading.Thread(target=self.run_thread, args=(procs,))
        thread.daemon = True
        thread.start()

    def stop_automation(self):
        if self.is_running:
            self.stop_flag = True
            self.status_var.set("Parando...")

    def run_thread(self, procs):
        automation.automacao_projudi(
            self.usuario_var.get(), self.senha_var.get(), self.oab_var.get(), procs,
            update_callback=self.update_ui,
            check_stop=lambda: self.stop_flag,
            log_callback=self.log_message ,
            lista_sucesso=self.sucessos
        )
        self.root.after(0, self.on_finished)

    def update_ui(self, msg, progress):
        self.root.after(0, lambda: self.status_var.set(msg))
        self.root.after(0, lambda: self.progress_var.set(progress))

    def on_finished(self):
            self.is_running = False
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            status = "Interrompido" if self.stop_flag else "Concluído"
            
            self.status_var.set(f"Execução: {status}")
            self.log_message(f"--- Fim da execução: {status} ---")

            if self.sucessos:
                try:
                    # Define o nome do arquivo Excel (mesma pasta do log, mas extensão .xlsx)
                    folder = os.path.dirname(self.current_log_file)
                    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                    excel_name = os.path.join(folder, f"Processos_Realizados_{timestamp}.xlsx")
                    
                    # Cria o DataFrame e salva
                    df_sucesso = pd.DataFrame(self.sucessos)
                    df_sucesso.to_excel(excel_name, index=False)
                    
                    self.log_message(f"✅ Relatório Excel gerado: {excel_name}")
                    msg_final = f"Relatório de Logs: {self.current_log_file}\n\n✅ Excel de Sucessos: {excel_name}"
                except Exception as e:
                    self.log_message(f"Erro ao gerar Excel: {e}")
                    msg_final = f"Relatório de Logs: {self.current_log_file}"
            else:
                self.log_message("Nenhum processo foi finalizado com sucesso para gerar o Excel.")
            
            # Avisa onde o arquivo foi salvo
            if self.current_log_file:
                self.log_message(f"✅ Logs salvos em: {self.current_log_file}")
                messagebox.showinfo("Relatório", f"Finalizado!\nRelatório salvo em:\n{excel_name}")
                
            self.stop_flag = False

    def log_message(self, message):
        # Usa 'after' para atualizar UI na thread principal
        self.root.after(0, lambda: self._append_log(message))
        # Grava no arquivo físico imediatamente
        self._save_to_file(message)

    def _append_log(self, message):
        # Adiciona hora na tela também
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {message}"
        
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, full_msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def _save_to_file(self, message):
        if self.current_log_file:
            try:
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with open(self.current_log_file, "a", encoding="utf-8") as f:
                    f.write(f"{timestamp} - {message}\n")
            except Exception as e:
                print(f"Erro ao salvar log em arquivo: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RPAApp(root)
    root.mainloop()
