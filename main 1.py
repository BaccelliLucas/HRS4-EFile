import sys
import requests
import customtkinter as ctk
from tkinter import filedialog
import threading
import os
from dotenv import load_dotenv
from functions import Functions
from openpyxl import load_workbook
from time import sleep
import logging
import getpass
import urllib3
import shutil

# --- [FIX PROXY BOSCH] PATCH SSL ---
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
_original_request = requests.Session.request
def _patched_request(self, method, url, *args, **kwargs):
    kwargs['verify'] = False
    return _original_request(self, method, url, *args, **kwargs)
requests.Session.request = _patched_request
# -----------------------------------

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

dotenv_path = os.path.join(base_path, '.env')
load_dotenv(dotenv_path)

user_pc = getpass.getuser()

# Configuração de Logs
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Download Arquivos e-File (Por EDV)")
        self.root.geometry("500x650")

        self.functions = Functions(self.root)
        self.collaborators_file_path = ""
        self.documents_file_path = ""
        self.benefits_file_path = ""
        self.contract_file_path = ""

        # --- GUI Elements ---
        self.select_file_frame = ctk.CTkFrame(root, fg_color="transparent")
        self.select_file_frame.place(relx=0.5, rely=0.15, anchor="center")
        ctk.CTkLabel(self.select_file_frame, text="1. Excel de Colaboradores:").pack()
        ctk.CTkButton(self.select_file_frame, text="Selecionar Arquivo", command = self.get_collaborator_file).pack()
        self.collaborators_path_label = ctk.CTkLabel(self.select_file_frame, wraplength=400, text="", text_color="gray")
        self.collaborators_path_label.pack()

        self.select_document_file_frame = ctk.CTkFrame(root, fg_color="transparent")
        self.select_document_file_frame.place(relx=0.5, rely=0.35, anchor="center") 
        ctk.CTkLabel(self.select_document_file_frame, text="2. Pasta de Documentos (Destino):").pack()
        ctk.CTkButton(self.select_document_file_frame, text="Selecionar Pasta", command=self.get_document_bot_folder).pack()
        self.documents_path_label = ctk.CTkLabel(self.select_document_file_frame, wraplength=400, text="", text_color="gray")
        self.documents_path_label.pack()

        self.select_benefits_file_frame = ctk.CTkFrame(root, fg_color="transparent")
        self.select_benefits_file_frame.place(relx=0.5, rely=0.55, anchor="center") 
        ctk.CTkLabel(self.select_benefits_file_frame, text="3. Pasta de Benefícios (Destino):").pack()
        ctk.CTkButton(self.select_benefits_file_frame, text="Selecionar Pasta", command=self.get_benenifts_bot_folder).pack()
        self.benefits_path_label = ctk.CTkLabel(self.select_benefits_file_frame, wraplength=400, text="", text_color="gray")
        self.benefits_path_label.pack()

        self.select_contract_file_frame = ctk.CTkFrame(root, fg_color="transparent")
        self.select_contract_file_frame.place(relx=0.5, rely=0.75, anchor="center") 
        ctk.CTkLabel(self.select_contract_file_frame, text="4. Pasta de Contratos (Destino):").pack()
        ctk.CTkButton(self.select_contract_file_frame, text="Selecionar Pasta", command=self.get_contract_bot_folder).pack()
        self.contract_path_label = ctk.CTkLabel(self.select_contract_file_frame, wraplength=400, text="", text_color="gray")
        self.contract_path_label.pack()

        self.start_button_frame = ctk.CTkFrame(root)
        self.start_button_frame.place(relx=0.5, rely=0.9, anchor="center")
        self.start_button = ctk.CTkButton(self.start_button_frame, text="INICIAR DOWNLOADS", width=200, height=40, command=self.start_process, state=ctk.DISABLED, fg_color="green")
        self.start_button.pack()

        self.loading_label = ctk.CTkLabel(root, text="", font=("Arial", 20))
        self.dots = ""
        self.loading = False
        self.error_label = ctk.CTkLabel(root, text="", font=("Arial", 12), text_color="red")

        self.DOWNLOAD_PATH = fr"C:\Users\{user_pc}\Documents\Download_documents_efile"

        self.workbook = None
        self.worksheet = None
        self.collaborator_list = []
        self.collaborators_error_list = []

        root.protocol("WM_DELETE_WINDOW", self.root.destroy)

    def verify_all_loaded_files(self):
        if all([self.collaborators_file_path, self.documents_file_path, self.benefits_file_path, self.contract_file_path]):
            self.start_button.configure(state=ctk.NORMAL)

    def get_collaborator_file(self):
        path = self.functions.select_file("Selecione Excel Colaboradores")
        if path:
            self.collaborators_file_path = path
            self.collaborators_path_label.configure(text=path)
            self.verify_all_loaded_files()

    def get_document_bot_folder(self):
        path = self.functions.select_folder("Pasta Documentos")
        if path:
            self.documents_file_path = path
            self.documents_path_label.configure(text=path)
            if not self.contract_file_path:
                self.contract_file_path = path
                self.contract_path_label.configure(text=path)
            self.verify_all_loaded_files()

    def get_benenifts_bot_folder(self):
        path = self.functions.select_folder("Pasta Benefícios")
        if path:
            self.benefits_file_path = path
            self.benefits_path_label.configure(text=path)
            self.verify_all_loaded_files()

    def get_contract_bot_folder(self):
        path = self.functions.select_folder("Pasta Contratos")
        if path:
            self.contract_file_path = path
            self.contract_path_label.configure(text=path)
            self.verify_all_loaded_files()

    def start_process(self):
        for frame in [self.select_file_frame, self.select_document_file_frame, self.select_benefits_file_frame, self.select_contract_file_frame, self.start_button_frame]:
            frame.destroy()

        self.loading_label.pack(pady=20)
        self.loading = True

        try:
            self.workbook = load_workbook(self.collaborators_file_path)
            self.worksheet = self.workbook.active
            self.collaborator_list = self.functions.get_collaborators_list(self.worksheet)
        except Exception as e:
            logging.error(f"Erro ao ler Excel: {e}")
            self.error_label.configure(text=f"Erro no Excel: {e}")
            self.error_label.pack()
            return

        if not self.collaborator_list:
            self.error_label.configure(text="Excel vazio ou formato errado!")
            self.error_label.pack()
            return

        self.animate()
        
        thread = threading.Thread(target=self.download_collaborators_files)
        thread.daemon = True
        thread.start()

    def animate(self):
        if self.loading:
            self.dots = "●" * ((len(self.dots) % 3) + 1)
            self.loading_label.configure(text=f"Processando{self.dots}")
            self.root.after(500, self.animate)

    def download_collaborators_files(self):
        logging.info("=========================================")
        logging.info("   INICIANDO DOWNLOADS E MESCLAGEM ")
        logging.info("=========================================")

        BASE_API = os.getenv("API_URL")
        headers = {
            "Authorization": os.getenv("AUTHORIZATION"),
            "Content-Type": "application/json"
        }

        total = len(self.collaborator_list)

        for index, collaborator in enumerate(self.collaborator_list):
            edv = collaborator.get("edv", "NA")
            cpf = collaborator.get("cpf", "")
            
            new_cpf = str(cpf).replace(".", "").replace("-", "")
            
            logging.info(f"\n[{index+1}/{total}] Processando EDV: {edv}")

            if len(new_cpf) != 11 or not new_cpf.isdigit():
                logging.warning(f"   [!] CPF Inválido ignorado: {cpf}")
                continue

            if self.functions.verify_ok_cell(self.worksheet, index+2):
                logging.info(f"   [i] Já processado. Pulando.")
                continue

            edv_folder = os.path.join(self.DOWNLOAD_PATH, str(edv))
            temp_docs_path = os.path.join(edv_folder, "Documents") 
            temp_benefits_path = os.path.join(edv_folder, "Benefits")

            os.makedirs(temp_docs_path, exist_ok=True)
            os.makedirs(temp_benefits_path, exist_ok=True)

            benefits_url = f"{BASE_API}/zip-benefits/{new_cpf}"
            documents_url = f"{BASE_API}/zip-documents/{new_cpf}"
            contract_url = f"{BASE_API}/work-contract/{new_cpf}"

            collaborator_error = False

            try:
                logging.info(f"   > Baixando Documentos...")
                self.functions.get_zip(url=documents_url, headers=headers, download_path=temp_docs_path, name="documents")
                
                logging.info(f"   > Baixando Contrato...")
                self.functions.get_pdf(url=contract_url, headers=headers, download_path=temp_docs_path, name="contract")

                logging.info(f"   > Baixando Benefícios...")
                self.functions.get_zip(url=benefits_url, headers=headers, download_path=temp_benefits_path, name="benefits")

            except Exception as e:
                logging.error(f"   [X] ERRO no download: {e}")
                self.collaborators_error_list.append(edv)
                collaborator_error = True
                continue

            if not collaborator_error:
                logging.info("   > Organizando arquivos nas pastas de destino final...")
                try:
                    # Cria as subpastas com o EDV dentro das pastas destino escolhidas no app
                    final_docs_dest = os.path.join(self.documents_file_path, str(edv))
                    final_benefits_dest = os.path.join(self.benefits_file_path, str(edv))
                    
                    os.makedirs(final_docs_dest, exist_ok=True)
                    os.makedirs(final_benefits_dest, exist_ok=True)

                    # -- Processa PASTA DOCUMENTOS E CONTRATO --
                    self.functions.get_file_in_zip(temp_docs_path)
                    sleep(1)
                    # Move para a subpasta do EDV no destino final mantendo o nome original
                    self.functions.move_files_to_BOT(temp_docs_path, final_docs_dest)

                    # === CHAMA A FUNÇÃO DE MESCLAGEM AQUI ===
                    logging.info("   > Mesclando documentos em um único PDF...")
                    self.functions.merge_files_to_pdf(final_docs_dest, output_filename="Contrato_Completo.pdf")

                    # -- Processa PASTA BENEFÍCIOS --
                    self.functions.get_file_in_zip(temp_benefits_path)
                    sleep(1)
                    # Move para a subpasta do EDV no destino final mantendo o nome original
                    self.functions.move_files_to_BOT(temp_benefits_path, final_benefits_dest)

                    logging.info("     [OK] Arquivos movidos e mesclados com sucesso.")

                except Exception as e:
                    logging.error(f"   [X] ERRO ao processar ou mesclar arquivos locais: {e}")
                    self.collaborators_error_list.append(edv)
                    collaborator_error = True

            if not collaborator_error:
                logging.info("   [SUCESSO] Finalizado.")
                self.functions.insert_ok_spreadsheet(self.worksheet, self.workbook, self.collaborators_file_path, index+2)
            else:
                logging.warning("   [FALHA] Marcando erro.")

        self.loading = False
        errors_qty = len(self.collaborators_error_list)
        
        logging.info("=========================================")
        logging.info(f"   FIM. Erros: {errors_qty}")
        logging.info("=========================================")

        if errors_qty > 0:
            msg = f"Finalizado com {errors_qty} erros!\nVerifique o log."
            self.error_label.configure(text=msg, text_color="red")
            self.functions.create_txt_in_file(self.collaborators_error_list, os.path.join(self.DOWNLOAD_PATH, "errors_log.txt"))
        else:
            self.error_label.configure(text="Sucesso Total!", text_color="green")
        
        self.error_label.pack(pady=10)
        self.loading_label.configure(text="Concluído")

if __name__ == "__main__":
    root = ctk.CTk()
    app = App(root)
    root.mainloop()