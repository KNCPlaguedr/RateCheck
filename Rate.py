# -*- coding: utf-8 -*-

import tkinter
from tkinter import ttk
import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import re
from datetime import datetime, date
import threading
import time

# Importações do Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

# --- IMPORTAÇÕES DE IA ---
try:
    from transformers import pipeline
    AI_AVAILABLE = True
except ImportError:
    AI_AVAILABLE = False

# --- CONFIGURAÇÕES ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# --- Mapeamento de Meses (Inglês -> Número) ---
MONTH_MAP = {
    'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
    'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'jun': 6, 'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
}

# --- CLASSE DE IA ---
class SmartEmailReader:
    def __init__(self):
        self.qa_pipeline = None
        self.model_loaded = False

    def load_model(self, update_callback=None):
        if not AI_AVAILABLE: return False, "Sem bibliotecas de IA"
        try:
            if update_callback: update_callback("Carregando motor IA (Backup)...")
            self.qa_pipeline = pipeline("question-answering", model="distilbert-base-cased-distilled-squad")
            self.model_loaded = True
            if update_callback: update_callback("Motor IA Pronto!")
            return True, "OK"
        except Exception as e:
            return False, str(e)

    def ask(self, context, question):
        if not self.model_loaded: return None
        return self.qa_pipeline(question=question, context=context[:3000])

ai_reader = SmartEmailReader()

# --- FUNÇÕES DE LIMPEZA E LÓGICA ---

def clean_money(val_str):
    if not isinstance(val_str, str): return 0.0
    # Remove espaços e símbolos de moeda, mantendo numeros, pontos e virgulas
    clean = re.sub(r'[^\d.,]', '', val_str)
    
    # Lógica para saber se ',' é decimal ou milhar
    if ',' in clean and '.' in clean: 
        # Ex: 1,842.98 -> tira a virgula
        clean = clean.replace(',', '') 
    elif ',' in clean:
        # Ex: 316,35 -> troca por ponto
        clean = clean.replace(',', '.')
    
    try:
        return float(clean)
    except:
        return 0.0

def parse_date_english(day_str, month_str, year_str):
    try:
        d = int(day_str)
        m = MONTH_MAP.get(month_str.lower(), 0)
        y = int(year_str)
        if m > 0:
            return datetime(y, m, d)
    except:
        pass
    return None

def find_rate_hybrid(email_text, target_date_str):
    """
    Lógica Avançada:
    1. Procura períodos de datas (from X to Y) e verifica se a Data Alvo está dentro.
    2. Se não achar por data, tenta o valor Total.
    3. Se não achar, tenta IA.
    """
    target_dt = None
    try:
        target_dt = datetime.strptime(target_date_str, "%d/%m/%Y")
    except:
        return 0.0, "Data Alvo Inválida"

    # --- ESTRATÉGIA 1: BUSCA POR PERÍODO (RATE CHANGES) ---
    # Procura linhas do tipo: "from Wednesday, November 19 2025 to Friday, November 21 2025 : R$401.40 BRL"
    # Regex flexível para capturar (Mes Dia Ano) do inicio e fim
    
    # Grupo 1: Mês Ini, G2: Dia Ini, G3: Ano Ini
    # Grupo 4: Mês Fim, G5: Dia Fim, G6: Ano Fim
    # Grupo 7: Valor
    pattern_period = re.compile(r"from .*?, ([A-Za-z]+) (\d+) (\d{4}) to .*?, ([A-Za-z]+) (\d+) (\d{4})\s*:\s*R\$\s*([\d.,]+)", re.IGNORECASE)
    
    matches = pattern_period.findall(email_text)
    
    for m in matches:
        month_i, day_i, year_i = m[0], m[1], m[2]
        month_f, day_f, year_f = m[3], m[4], m[5]
        rate_str = m[6]
        
        dt_start = parse_date_english(day_i, month_i, year_i)
        dt_end = parse_date_english(day_f, month_f, year_f)
        
        if dt_start and dt_end:
            # Verifica se a data alvo está no intervalo [Inicio, Fim)
            if dt_start <= target_dt < dt_end:
                return clean_money(rate_str), f"Tarifa do Período ({day_i}/{month_i}-{day_f}/{month_f})"

    # --- ESTRATÉGIA 2: PADRÃO ACCOR GENÉRICO (Se não achou período específico) ---
    accor_simple = re.search(r":\s*R\$\s*([\d.,]+)\s*BRL\s*per night", email_text, re.IGNORECASE)
    if accor_simple and not matches: 
        return clean_money(accor_simple.group(1)), "Padrão Accor (Único)"

    # --- ESTRATÉGIA 4: IA (BACKUP FINAL) ---
    if ai_reader.model_loaded:
        res = ai_reader.ask(email_text, f"What is the daily rate for {target_date_str}?")
        if res and res['score'] > 0.1:
            txt = res['answer']
            val = clean_money(txt)
            if 10 < val < 5000:
                return val, f"IA ({res['score']:.2f})"

    return 0.0, "Não encontrado"

# --- PROCESSAMENTO (WORKER) ---
def process_reservations(csv_paths, target_date_str, ignore_set, update_log_callback, update_progress_callback, on_complete_callback, stop_event):
    results_data = []
    verified_correct = []
    no_reference_rooms = []
    incorrect_rate_rooms = []
    
    try:
        update_log_callback("Inicializando IA e Navegador...")
        ai_reader.load_model(update_log_callback)

        options = webdriver.ChromeOptions()
        options.debugger_address = "localhost:9222"
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

        # Leitura Robusta dos CSVs
        all_dfs = []
        for path in csv_paths:
            try:
                df_temp = pd.read_csv(path, dtype=str).fillna('')
                all_dfs.append(df_temp)
            except Exception as e:
                update_log_callback(f"Erro ao ler arquivo {path}: {e}")

        if not all_dfs:
            update_log_callback("Nenhum CSV carregado corretamente.")
            on_complete_callback([], [], [], [])
            return

        df = pd.concat(all_dfs, ignore_index=True)
        total_rows = len(df)
        update_log_callback(f"Total de reservas para analisar: {total_rows}")

        for index, row in df.iterrows():
            if stop_event.is_set():
                update_log_callback("--- INTERROMPIDO ---")
                break

            progress_value = (index + 1) / total_rows
            percent = int(progress_value * 100)
            update_progress_callback(progress_value, f"Processando {index + 1} de {total_rows} ({percent}%)")

            # Dados com tratamento de erro
            try:
                ext_ref = str(row.get("External Reference", "")).strip()
                rate_csv_str = str(row.get("Rate", "")).strip()
                name = str(row.get("Name", "N/A")).strip()
                room = str(row.get("Room", "N/A")).strip()
                adults_val = row.get("Adults", "0")
            except Exception as e:
                update_log_callback(f"Erro ao ler linha {index}: {e}")
                continue

            # Lógica SHARE (Adultos < 1)
            try:
                adults_count = int(float(adults_val)) if adults_val else 0
            except:
                adults_count = 0

            if adults_count < 1:
                 results_data.append({'Quarto': room, 'Nome': name, 'Ref.': ext_ref, 'Tarifa CSV': rate_csv_str, 'Tarifa Email': '', 'Status': 'IGNORADO (SHARE)'})
                 continue
            
            if not ext_ref:
                no_reference_rooms.append(room)
                results_data.append({'Quarto': room, 'Nome': name, 'Ref.': ext_ref, 'Tarifa CSV': rate_csv_str, 'Tarifa Email': '', 'Status': 'SEM REF.'})
                continue
            
            if room in ignore_set:
                results_data.append({'Quarto': room, 'Nome': name, 'Ref.': ext_ref, 'Tarifa CSV': rate_csv_str, 'Tarifa Email': '', 'Status': 'IGNORADO (QUARTO)'})
                continue

            update_log_callback(f"--- Quarto: {room} (Ref: {ext_ref}) ---")

            try:
                # Busca
                try:
                    search_box = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, 'topSearchInput')))
                    search_box.clear(); search_box.send_keys(Keys.CONTROL + 'a'); search_box.send_keys(Keys.DELETE)
                    search_box.send_keys(ext_ref); search_box.send_keys(Keys.ENTER)
                except:
                    update_log_callback("Erro: Barra de busca inacessível.")
                    results_data.append({'Quarto': room, 'Nome': name, 'Ref.': ext_ref, 'Tarifa CSV': rate_csv_str, 'Tarifa Email': 'ERRO', 'Status': 'ERRO BUSCA'})
                    continue
                
                # Clica no Resultado
                try:
                    email_result = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f'//div[@role="option" and contains(@aria-label, "{ext_ref}")]')))
                    email_result.click()
                except:
                    update_log_callback("E-mail não encontrado.")
                    results_data.append({'Quarto': room, 'Nome': name, 'Ref.': ext_ref, 'Tarifa CSV': rate_csv_str, 'Tarifa Email': 'N/A', 'Status': 'EMAIL NÃO ENCONTRADO'})
                    continue

                # --- ATUALIZAÇÃO PARA LER CORPO (CORREÇÃO APLICADA) ---
                try:
                    # TENTATIVA 1: Pelo ID específico que você forneceu
                    try:
                        email_body = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.ID, 'Pular para mensagem-region'))
                        )
                    except:
                        # TENTATIVA 2: Pelo role="main" (caso o ID mude, mas a estrutura se mantenha)
                        email_body = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="main"]'))
                        )

                    time.sleep(0.5)
                    email_text = email_body.text
                except Exception as ex_body:
                    update_log_callback(f"Erro ao ler texto do email: {ex_body}")
                    results_data.append({'Quarto': room, 'Nome': name, 'Ref.': ext_ref, 'Tarifa CSV': rate_csv_str, 'Tarifa Email': '', 'Status': 'ERRO LEITURA'})
                    continue
                
                # Tenta fechar
                try: driver.find_element(By.CSS_SELECTOR, 'button[title="Fechar"]').click()
                except: pass

                # --- LÓGICA HÍBRIDA ---
                rate_email_val, method_msg = find_rate_hybrid(email_text, target_date_str)
                rate_csv_val = clean_money(rate_csv_str)

                # Comparação
                if abs(rate_csv_val - rate_email_val) < 1.00 and rate_email_val > 0:
                    status = 'CORRETO'
                    verified_correct.append(room)
                else:
                    status = 'ERRO DE TARIFA'
                    incorrect_rate_rooms.append(room)

                update_log_callback(f"Status: {status} | CSV: {rate_csv_val} | Email: {rate_email_val} [{method_msg}]")
                results_data.append({'Quarto': room, 'Nome': name, 'Ref.': ext_ref, 'Tarifa CSV': f"R${rate_csv_val:.2f}", 'Tarifa Email': f"R${rate_email_val:.2f}", 'Status': status})

            except Exception as e:
                update_log_callback(f"Erro Genérico no Quarto {room}: {str(e)}")
                results_data.append({'Quarto': room, 'Nome': name, 'Ref.': ext_ref, 'Tarifa CSV': rate_csv_str, 'Tarifa Email': 'ERRO', 'Status': 'ERRO GERAL'})
                continue

    except Exception as e:
        update_log_callback(f"ERRO CRÍTICO GLOBAL: {str(e)}")
    finally:
        update_log_callback("--- FIM ---")
        on_complete_callback(results_data, verified_correct, no_reference_rooms, incorrect_rate_rooms)

# --- INTERFACE GRÁFICA ---
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Verificador de Tarifas (Fix Leitura)")
        self.geometry("900x650")
        
        self.csv_paths = []
        self.session_results = {}
        self.stop_event = threading.Event()

        # Layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # Topo
        self.top_frame = ctk.CTkFrame(self)
        self.top_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.top_frame.grid_columnconfigure(1, weight=1)

        self.btn_select_csv = ctk.CTkButton(self.top_frame, text="Selecionar CSV(s)", command=self.select_csv_files)
        self.btn_select_csv.grid(row=0, column=0, padx=10, pady=10)

        self.lbl_file_path = ctk.CTkLabel(self.top_frame, text="Nenhum arquivo selecionado")
        self.lbl_file_path.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # Inputs
        self.input_frame = ctk.CTkFrame(self)
        self.input_frame.grid(row=1, column=0, padx=10, pady=0, sticky="ew")
        self.input_frame.grid_columnconfigure(1, weight=1) 

        ctk.CTkLabel(self.input_frame, text="Data Alvo (DD/MM/AAAA):").grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")
        self.date_entry = ctk.CTkEntry(self.input_frame)
        self.date_entry.grid(row=0, column=1, padx=10, pady=(10, 5), sticky="ew")
        self.date_entry.insert(0, date.today().strftime("%d/%m/%Y"))

        ctk.CTkLabel(self.input_frame, text="Ignorar Quartos (ex: 101, 102):").grid(row=1, column=0, columnspan=2, padx=10, pady=(5,0), sticky="w")
        self.ignore_textbox = ctk.CTkTextbox(self.input_frame, height=40)
        self.ignore_textbox.grid(row=2, column=0, columnspan=2, padx=10, pady=(0,10), sticky="ew")

        # Ações
        self.action_frame = ctk.CTkFrame(self)
        self.action_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        self.action_frame.grid_columnconfigure((0, 1, 2), weight=1)

        self.btn_start = ctk.CTkButton(self.action_frame, text="Iniciar Verificação", command=self.start_processing_thread, state="disabled")
        self.btn_start.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.btn_stop = ctk.CTkButton(self.action_frame, text="Interromper", command=self.stop_processing, state="disabled", fg_color="#D32F2F", hover_color="#B71C1C")
        self.btn_stop.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.btn_show_verified = ctk.CTkButton(self.action_frame, text="Ver Resumo", command=self.show_summary_window, state="disabled")
        self.btn_show_verified.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

        # Abas
        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.grid(row=3, column=0, padx=10, pady=0, sticky="nsew")
        self.tab_view.add("Processo"); self.tab_view.add("Relatório Final")

        # Aba Processo
        process_tab_frame = self.tab_view.tab("Processo")
        process_tab_frame.grid_columnconfigure(0, weight=1); process_tab_frame.grid_rowconfigure(1, weight=1)
        
        progress_frame = ctk.CTkFrame(process_tab_frame, fg_color="transparent")
        progress_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        progress_frame.grid_columnconfigure(0, weight=1)
        
        self.progress_bar = ctk.CTkProgressBar(progress_frame); self.progress_bar.set(0)
        self.progress_bar.grid(row=0, column=0, padx=5, pady=2, sticky="ew")
        self.progress_label = ctk.CTkLabel(progress_frame, text="Aguardando...")
        self.progress_label.grid(row=0, column=1, padx=5, pady=2)
        
        self.log_textbox = ctk.CTkTextbox(process_tab_frame, state="disabled")
        self.log_textbox.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")

        # Aba Relatório
        report_frame = ctk.CTkFrame(self.tab_view.tab("Relatório Final"), fg_color="transparent")
        report_frame.pack(expand=True, fill="both", padx=5, pady=5)
        self.create_report_table(report_frame)

    def create_report_table(self, parent):
        style = ttk.Style(); style.theme_use("default")
        style.configure("Treeview", background="#2a2d2e", foreground="white", fieldbackground="#343638", borderwidth=0, rowheight=25)
        style.map('Treeview', background=[('selected', '#24527d')])
        style.configure("Treeview.Heading", background="#565b5e", foreground="white", relief="flat", font=('Segoe UI', 9, 'bold'))
        style.map("Treeview.Heading", background=[('active', '#343638')])
        
        self.report_tree = ttk.Treeview(parent, style="Treeview")
        self.report_tree['columns'] = ('Quarto', 'Nome', 'Ref.', 'Tarifa CSV', 'Tarifa Email', 'Status')
        self.report_tree.column("#0", width=0, stretch=False); self.report_tree.column("Quarto", anchor='center', width=60)
        self.report_tree.column("Nome", anchor='w', width=200); self.report_tree.column("Ref.", anchor='w', width=120)
        self.report_tree.column("Tarifa CSV", anchor='e', width=100); self.report_tree.column("Tarifa Email", anchor='e', width=100)
        self.report_tree.column("Status", anchor='w', width=150)
        self.report_tree.heading("Quarto", text="Quarto"); self.report_tree.heading("Nome", text="Nome"); self.report_tree.heading("Ref.", text="Ref.")
        self.report_tree.heading("Tarifa CSV", text="Tarifa CSV"); self.report_tree.heading("Tarifa Email", text="Tarifa Email"); self.report_tree.heading("Status", text="Status")
        
        # Tags coloridas
        self.report_tree.tag_configure('correct', foreground='#66FF66')
        self.report_tree.tag_configure('error', foreground='#FF4D4D')
        self.report_tree.tag_configure('warning', foreground='#FFCC00')
        self.report_tree.tag_configure('share', foreground='#33CCFF')
        self.report_tree.tag_configure('ignored', foreground='#9E9E9E')
        
        self.report_tree.pack(expand=True, fill="both")

    def get_tag_for_status(self, status):
        status_str = str(status).upper()
        if 'CORRETO' in status_str: return 'correct'
        if 'ERRO DE TARIFA' in status_str: return 'error'
        if 'SHARE' in status_str: return 'share'
        if 'IGNORADO' in status_str: return 'ignored'
        return 'warning'

    def populate_report_tab(self, data):
        for item in self.report_tree.get_children(): self.report_tree.delete(item)
        for record in data:
            tag = self.get_tag_for_status(record.get('Status', ''))
            self.report_tree.insert(parent='', index='end', values=list(record.values()), tags=(tag,))

    def on_processing_complete(self, results, correct_list, no_ref_list, wrong_rate_list):
        self.populate_report_tab(results)
        self.session_results = {'correct': correct_list, 'no_ref': no_ref_list, 'wrong_rate': wrong_rate_list}
        self.btn_start.configure(state="normal", text="Iniciar Verificação"); self.btn_stop.configure(state="disabled", text="Interromper")
        if any(self.session_results.values()): self.btn_show_verified.configure(state="normal")
        self.tab_view.set("Relatório Final")
        if not self.stop_event.is_set(): messagebox.showinfo("Concluído", "Verificação Finalizada!")
        else: messagebox.showwarning("Interrompido", "Parado pelo usuário.")

    def show_summary_window(self):
        summary_win = ctk.CTkToplevel(self); summary_win.title("Resumo"); summary_win.geometry("450x400")
        ctk.CTkLabel(summary_win, text="✅ Quartos Corretos:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10,2))
        t1 = ctk.CTkTextbox(summary_win, height=80); t1.pack(fill="x", padx=10); t1.insert("1.0", ",".join(self.session_results.get('correct',[])))
        ctk.CTkLabel(summary_win, text="⚠️ Sem Referência:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10,2))
        t2 = ctk.CTkTextbox(summary_win, height=80); t2.pack(fill="x", padx=10); t2.insert("1.0", ",".join(self.session_results.get('no_ref',[])))
        ctk.CTkLabel(summary_win, text="❌ Rate Incorreto:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10,2))
        t3 = ctk.CTkTextbox(summary_win, height=80); t3.pack(fill="x", padx=10); t3.insert("1.0", ",".join(self.session_results.get('wrong_rate',[])))

    def select_csv_files(self):
        self.csv_paths = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
        if self.csv_paths:
            self.lbl_file_path.configure(text=f"{len(self.csv_paths)} arquivo(s) selecionado(s)")
            self.btn_start.configure(state="normal")

    def update_log(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_textbox.configure(state="disabled"); self.log_textbox.see("end")

    def update_progress(self, value, text):
        self.progress_bar.set(value); self.progress_label.configure(text=text)
    
    def stop_processing(self):
        self.stop_event.set(); self.btn_stop.configure(state="disabled", text="Parando...")

    def start_processing_thread(self):
        if not self.date_entry.get().strip(): return
        self.stop_event.clear()
        self.btn_start.configure(state="disabled", text="Processando..."); self.btn_stop.configure(state="normal")
        self.log_textbox.configure(state="normal"); self.log_textbox.delete("1.0", "end"); self.log_textbox.configure(state="disabled")
        
        ignore_set = {x.strip() for x in self.ignore_textbox.get("1.0", "end").split(',') if x.strip()}
        
        thread = threading.Thread(target=process_reservations, args=(self.csv_paths, self.date_entry.get().strip(), ignore_set, self.update_log, self.update_progress, lambda r, c, n, w: self.after(0, self.on_processing_complete, r, c, n, w), self.stop_event))
        thread.daemon = True
        thread.start()

if __name__ == "__main__":
    app = App()
    app.mainloop()