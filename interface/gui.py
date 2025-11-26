import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import pandas as pd
import time
import os
import json
import queue
from core.processor import processar_contratos
from utils.config_handler import salvar_config, carregar_config
from utils.validador import validar_planilha
from automation.bot import controle_parada

def iniciar_interface():
    class App(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title("🤖 Automação CREA-MG - CREA-BOT")

            self.config = carregar_config()
            ctk.set_appearance_mode(self.config.get("tema", "Dark"))

            # Recupera tamanho e posição salvos (ou usa padrão)
            pos = self.config.get("posicao_janela")
            if pos and isinstance(pos, dict):
                try:
                    self.geometry(f"{pos.get('width', 1000)}x{pos.get('height', 600)}+{pos.get('x', 100)}+{pos.get('y', 100)}")
                except:
                    self.geometry("1000x600")
            else:
                self.geometry("1000x600")

            self.resizable(True, True)  # Libera resize!

            self.planilha = None
            self.processando = False
            self.fila_logs = queue.Queue()

            self.toast = None  # <-- Para notificações toast

            self.create_sidebar()
            self.create_main_area()
            self.atualizar_texto_tema()

            self.carregar_ultima_planilha()  # <-- Carrega a última planilha ao abrir

            self.bind("<Configure>", self.salvar_posicao_janela)
            self.protocol("WM_DELETE_WINDOW", self.fechar_app)
            self.after(100, self.checar_fila_logs)

        def show_toast(self, msg, duration=2000):  # 2 segundos padrão
            # Evita várias toasts ao mesmo tempo
            if self.toast:
                self.toast.destroy()
            self.toast = ctk.CTkLabel(self, text=msg, fg_color="black", text_color="white", font=("Arial", 14, "bold"))
            self.toast.place(relx=0.5, rely=0.05, anchor="n")
            self.toast.after(duration, self.toast.destroy)

        def carregar_ultima_planilha(self):
            caminho_anterior = self.config.get("caminho_planilha")
            if caminho_anterior and os.path.exists(caminho_anterior):
                try:
                    engine = "openpyxl" if caminho_anterior.endswith(".xlsx") else None
                    self.planilha = pd.read_excel(
                        caminho_anterior,
                        dtype=str,
                        engine=engine
                    )
                    self.label_planilha.configure(text=f"Planilha carregada: {os.path.basename(caminho_anterior)}")
                    self.btn_iniciar.configure(state="normal")
                    self.log("✅ Última planilha carregada automaticamente.")
                    self.show_toast("Última planilha carregada automaticamente!")

                    ultimo_contrato_real = self.config.get("ultimo_contrato")
                    if ultimo_contrato_real:
                        try:
                            proximo = str(int(ultimo_contrato_real) + 1)
                            self.entry_inicio.delete(0, "end")
                            self.entry_inicio.insert(0, proximo)
                            self.log(f"🔢 Sugerido iniciar no contrato {proximo}")
                        except Exception as e:
                            self.log(f"⚠️ Erro ao sugerir próximo contrato: {e}")

                except Exception as e:
                    self.log(f"⚠️ Erro ao carregar última planilha: {e}")

        def salvar_posicao_janela(self, event=None):
            if event and event.widget == self:
                self.config["posicao_janela"] = {
                    "x": self.winfo_x(),
                    "y": self.winfo_y(),
                    "width": self.winfo_width(),
                    "height": self.winfo_height()
                }
                salvar_config(self.config)

        def create_sidebar(self):
            self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
            self.sidebar.pack(side="left", fill="y")

            self.label_logo = ctk.CTkLabel(self.sidebar, text="🤖 CREA-BOT", font=("Arial", 24, "bold"))
            self.label_logo.pack(pady=(30, 20))

            self.btn_carregar = ctk.CTkButton(self.sidebar, text="📂 Carregar Planilha", command=self.carregar_planilha)
            self.btn_carregar.pack(pady=10, fill="x")

            self.entry_inicio = ctk.CTkEntry(self.sidebar, placeholder_text="Número do Contrato inicial")
            self.entry_inicio.pack(pady=(30, 10), fill="x")

            self.btn_iniciar = ctk.CTkButton(self.sidebar, text="▶ Iniciar", command=self.iniciar_bot, state="disabled")
            self.btn_iniciar.pack(pady=10, fill="x")

            self.btn_parar = ctk.CTkButton(self.sidebar, text="⏹ Parar", command=self.parar_bot, state="disabled")
            self.btn_parar.pack(pady=10, fill="x")

            self.btn_fechar = ctk.CTkButton(self.sidebar, text="❌ Fechar App", fg_color="red", hover_color="#990000", command=self.fechar_app)
            self.btn_fechar.pack(pady=(40, 10), fill="x")

            self.btn_mudar_tema = ctk.CTkButton(self.sidebar, command=self.alternar_tema)
            self.btn_mudar_tema.pack(pady=10, fill="x")
            self.btn_mudar_tema.bind("<Enter>", lambda e: self.mostrar_tooltip("Clique para alternar tema Claro/Escuro"))
            self.btn_mudar_tema.bind("<Leave>", lambda e: self.esconder_tooltip())
            self.tooltip = None

        def create_main_area(self):
            self.main_area = ctk.CTkFrame(self)
            self.main_area.pack(expand=True, fill="both", padx=10, pady=10)

            self.label_planilha = ctk.CTkLabel(self.main_area, text="Nenhuma planilha carregada.", font=("Arial", 14))
            self.label_planilha.pack(pady=10)

            self.progress_bar = ctk.CTkProgressBar(self.main_area)
            self.progress_bar.pack(fill="x", padx=20, pady=10)
            self.progress_bar.set(0)

            # Barra de progresso detalhada (percentual)
            self.progress_percent = ctk.CTkLabel(self.main_area, text="0%", font=("Arial", 14, "bold"))
            self.progress_percent.pack(pady=(0, 10))

            self.log_textbox = ctk.CTkTextbox(self.main_area, height=400)
            self.log_textbox.pack(fill="both", expand=True, padx=20, pady=10)
            self.log_textbox.configure(state="disabled")

        def mostrar_tooltip(self, texto):
            self.tooltip = ctk.CTkLabel(self.btn_mudar_tema, text=texto, font=("Arial", 10), bg_color="gray")
            self.tooltip.place(relx=1.05, rely=0.5, anchor="w")

        def esconder_tooltip(self):
            if self.tooltip:
                self.tooltip.destroy()
                self.tooltip = None

        def alternar_tema(self):
            atual = ctk.get_appearance_mode()
            novo = "Dark" if atual == "Light" else "Light"
            ctk.set_appearance_mode(novo)
            self.config["tema"] = novo
            salvar_config(self.config)
            self.atualizar_texto_tema()

        def atualizar_texto_tema(self):
            tema = ctk.get_appearance_mode()
            self.btn_mudar_tema.configure(text="🌞 Tema Claro" if tema == "Dark" else "🌙 Tema Escuro")

        def carregar_planilha(self):
            caminho = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx *.xlsm")])
            if not caminho:
                return
            try:
                engine = "openpyxl" if caminho.endswith(".xlsx") else None
                df = pd.read_excel(
                    caminho,
                    dtype=str,
                    engine=engine
                )

                erros = validar_planilha(df)
                if erros:
                    messagebox.showerror("Erros na planilha", "\n".join(erros))
                    self.log(f"❌ Erros na planilha:\n" + "\n".join(erros))
                    return

                self.planilha = df
                self.label_planilha.configure(text=f"Planilha carregada: {os.path.basename(caminho)}")
                self.btn_iniciar.configure(state="normal")
                self.config["caminho_planilha"] = caminho
                salvar_config(self.config)
                self.log("✅ Planilha carregada com sucesso!")
                self.show_toast("Planilha carregada com sucesso!")

                ultimo_contrato_real = self.config.get("ultimo_contrato")
                if ultimo_contrato_real:
                    try:
                        proximo = str(int(ultimo_contrato_real) + 1)
                        self.entry_inicio.delete(0, "end")
                        self.entry_inicio.insert(0, proximo)
                    except ValueError:
                        self.entry_inicio.delete(0, "end")
                        self.entry_inicio.insert(0, "")
                else:
                    self.entry_inicio.delete(0, "end")
                    self.entry_inicio.insert(0, "")

            except Exception as e:
                messagebox.showerror("Erro ao carregar planilha", str(e))
                self.log(f"❌ Erro ao carregar planilha: {e}")

        def atualizar_numero_contrato(self, numero):  # 🔥 NOVO
            self.config["ultimo_contrato"] = str(numero)
            salvar_config(self.config)
            try:
                proximo = str(int(numero) + 1)
                self.entry_inicio.delete(0, "end")
                self.entry_inicio.insert(0, proximo)
                self.log(f"🔢 Próximo contrato sugerido: {proximo}")
            except Exception as e:
                self.log(f"⚠️ Erro ao sugerir próximo contrato: {e}")

        def iniciar_bot(self):
            if self.processando or self.planilha is None:
                return

            numero_inicial = self.entry_inicio.get().strip()
            if not numero_inicial.isdigit():
                messagebox.showerror("Erro", "Digite um número de contrato válido.")
                return

            linha_match = self.planilha[self.planilha["NUMERO DO CONTRATO"].astype(str) == numero_inicial]
            if linha_match.empty:
                messagebox.showerror("Erro", f"Número {numero_inicial} não encontrado na planilha.")
                return

            cpf_cnpj_inicial = linha_match["CPF_CNPJ"].iloc[0]
            data_registro_inicial = linha_match["DATA DO REGISTRO"].iloc[0]

            self.processando = True
            self.btn_iniciar.configure(state="disabled")
            self.btn_parar.configure(state="normal")
            self.progress_bar.set(0)
            self.progress_percent.configure(text="0%")

            # 🔥 NOVO: Callback para atualizar o número do contrato
            def callback_atualizar_contrato(numero):
                self.atualizar_numero_contrato(numero)

            thread = threading.Thread(
                target=processar_contratos,
                args=(
                    self.planilha,
                    0,
                    self.log,
                    numero_inicial,
                    cpf_cnpj_inicial,
                    data_registro_inicial,
                    callback_atualizar_contrato  # 🔥 NOVO!
                ),
                daemon=True
            )
            thread.start()

        def parar_bot(self):
            controle_parada["parar"] = True
            self.log("🛑 Parada solicitada. Aguardando finalização da tarefa atual...")
            self.btn_parar.configure(state="disabled")

        def log(self, msg):
            self.fila_logs.put(msg)

        def checar_fila_logs(self):
            while not self.fila_logs.empty():
                msg = self.fila_logs.get()
                self.log_textbox.configure(state="normal")
                self.log_textbox.insert("end", msg + "\n")
                self.log_textbox.see("end")
                self.log_textbox.configure(state="disabled")
                if msg.startswith("Progresso:"):
                    try:
                        valor = float(msg.split(":")[1].replace("%", "").strip()) / 100
                        self.progress_bar.set(valor)
                        self.progress_percent.configure(text=f"{int(valor*100)}%")
                    except:
                        self.progress_percent.configure(text="0%")
            self.after(200, self.checar_fila_logs)

        def fechar_app(self):
            if self.processando:
                resposta = messagebox.askyesno(
                    "Confirmar saída",
                    "O processamento está em andamento.\nTem certeza que deseja fechar o app?"
                )
                if not resposta:
                    return
            else:
                resposta = messagebox.askyesno("Confirmar saída", "Tem certeza que deseja fechar o app?")
                if not resposta:
                    return
            salvar_config(self.config)
            self.destroy()

    ctk.set_default_color_theme("blue")
    app = App()
    app.mainloop()
