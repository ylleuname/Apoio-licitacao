import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os

from extrator_pncp import buscar_editais_pncp, salvar_em_csv

# --- CONFIGURAÇÃO DA PALETA DE CORES BLOKO ---
COR_PRIMARIA = "#3751A3"
COR_SECUNDARIA = "#a38937"
COR_CLARA = "#7188cf"
COR_FUNDO = "#FFFFFF"
COR_CARD = "#F8F9FA"

ctk.set_appearance_mode("light")


class AppMonitoramento(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Bloko - Monitoramento de Editais")
        self.geometry("700x650")  # Tamanho fixo e confortável
        self.resizable(False, False)
        self.configure(fg_color=COR_FUNDO)

        # --- 1. HEADER (Fixo no topo) ---
        self.frame_header = ctk.CTkFrame(self, fg_color=COR_PRIMARIA, corner_radius=0, height=80)
        self.frame_header.pack(fill="x", side="top")
        self.frame_header.pack_propagate(False)

        self.lbl_titulo = ctk.CTkLabel(
            self.frame_header,
            text="Inteligência de Licitações - PNCP",
            font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
            text_color="#FFFFFF"
        )
        self.lbl_titulo.pack(pady=22)

        # --- 2. FOOTER (Fixo na base da janela) ---
        self.frame_footer = ctk.CTkFrame(self, fg_color=COR_FUNDO)
        self.frame_footer.pack(fill="x", side="bottom", padx=35, pady=(10, 15))

        self.btn_iniciar = ctk.CTkButton(
            self.frame_footer,
            text="INICIAR MONITORAMENTO",
            command=self.iniciar_busca,
            fg_color=COR_SECUNDARIA,
            hover_color="#8c752d",
            text_color="#FFFFFF",
            height=50,
            font=ctk.CTkFont(size=15, weight="bold"),
            corner_radius=8
        )
        self.btn_iniciar.pack(fill="x", pady=(0, 5))

        self.lbl_status = ctk.CTkLabel(
            self.frame_footer,
            text="Pronto para iniciar.",
            text_color="#777777",
            font=ctk.CTkFont(size=12)
        )
        self.lbl_status.pack(pady=0)

        # --- 3. CORPO COM ROLAGEM ---
        self.container = ctk.CTkScrollableFrame(
            self,
            fg_color=COR_FUNDO,
            scrollbar_button_color=COR_CLARA,
            scrollbar_button_hover_color=COR_PRIMARIA
        )
        self.container.pack(fill="both", expand=True, padx=35, pady=(20, 0))

        # Filtros Base
        self.lbl_secao1 = ctk.CTkLabel(self.container, text="Filtros Base (Catálogo)",
                                       font=ctk.CTkFont(size=14, weight="bold"), text_color=COR_PRIMARIA)
        self.lbl_secao1.pack(anchor="w", pady=(0, 5))

        self.frame_checks = ctk.CTkFrame(self.container, fg_color=COR_CARD, border_width=1, border_color="#EBEBEB",
                                         corner_radius=8)
        self.frame_checks.pack(fill="x", pady=(0, 20), ipadx=10, ipady=10)

        self.frame_checks.grid_columnconfigure(0, weight=1)
        self.frame_checks.grid_columnconfigure(1, weight=1)

        self.check_vars = {}

        termos_brutos = [
            "videowall", "surveillance", "controle de acesso",
            "infraestrutura seca", "cabeamento estruturado",
            "fibra óptica", "data center", "access point", "switch"
        ]

        termos_padrao = sorted([termo.title() for termo in termos_brutos])

        for i, termo in enumerate(termos_padrao):
            var = ctk.BooleanVar(value=True)
            chk = ctk.CTkCheckBox(
                self.frame_checks,
                text=termo,
                variable=var,
                font=ctk.CTkFont(size=13),
                fg_color=COR_PRIMARIA,
                hover_color=COR_CLARA,
                text_color="#333333",
                border_color=COR_PRIMARIA,
                checkbox_width=22,
                checkbox_height=22
            )
            linha = i // 2
            coluna = i % 2
            chk.grid(row=linha, column=coluna, padx=20, pady=10, sticky="w")
            self.check_vars[termo] = var

        # Termos Adicionais
        self.lbl_secao2 = ctk.CTkLabel(self.container, text="Termos Adicionais",
                                       font=ctk.CTkFont(size=14, weight="bold"), text_color=COR_PRIMARIA)
        self.lbl_secao2.pack(anchor="w", pady=(0, 5))

        self.entry_custom = ctk.CTkEntry(
            self.container,
            placeholder_text="Insira novos termos separados por vírgula (Ex: router, firewall, gpon)",
            height=45,
            font=ctk.CTkFont(size=13),
            border_color="#CCCCCC",
            fg_color="#FFFFFF",
            text_color="#333333",
            corner_radius=6
        )
        self.entry_custom.pack(fill="x", pady=(0, 20))

        # Configurações de Extração
        self.lbl_secao_download = ctk.CTkLabel(self.container, text="Configurações de Extração",
                                               font=ctk.CTkFont(size=14, weight="bold"), text_color=COR_PRIMARIA)
        self.lbl_secao_download.pack(anchor="w", pady=(0, 5))

        self.switch_baixar = ctk.CTkSwitch(
            self.container,
            text="Baixar anexos e editais em PDF",
            font=ctk.CTkFont(size=13),
            text_color="#333333",
            progress_color=COR_PRIMARIA,
            button_color=COR_SECUNDARIA,
            button_hover_color="#8c752d"
        )
        self.switch_baixar.select()
        self.switch_baixar.pack(anchor="w", pady=(0, 20))

        # Local de Armazenamento
        self.lbl_secao3 = ctk.CTkLabel(self.container, text="Local de Armazenamento",
                                       font=ctk.CTkFont(size=14, weight="bold"), text_color=COR_PRIMARIA)
        self.lbl_secao3.pack(anchor="w", pady=(0, 5))

        self.frame_dir = ctk.CTkFrame(self.container, fg_color="transparent")
        self.frame_dir.pack(fill="x", pady=(0, 10))

        self.caminho_salvamento = os.path.join(os.path.expanduser("~"), "Desktop", "Editais_PNCP")

        self.entry_dir = ctk.CTkEntry(
            self.frame_dir,
            height=40,
            font=ctk.CTkFont(size=13),
            text_color="#666666",
            fg_color="#F2F2F2",
            border_width=1,
            border_color="#CCCCCC",
            corner_radius=6
        )
        self.entry_dir.pack(side="left", fill="x", expand=True, padx=(0, 15))
        self.entry_dir.insert(0, self.caminho_salvamento)
        self.entry_dir.configure(state="disabled")

        self.btn_dir = ctk.CTkButton(
            self.frame_dir,
            text="Alterar Pasta",
            command=self.selecionar_diretorio,
            fg_color=COR_PRIMARIA,
            hover_color=COR_CLARA,
            text_color="#FFFFFF",
            font=ctk.CTkFont(weight="bold"),
            height=40,
            width=130,
            corner_radius=6
        )
        self.btn_dir.pack(side="right")

    def selecionar_diretorio(self):
        diretorio = filedialog.askdirectory(title="Selecione a pasta para guardar os arquivos")
        if diretorio:
            self.caminho_salvamento = diretorio
            self.entry_dir.configure(state="normal")
            self.entry_dir.delete(0, "end")
            self.entry_dir.insert(0, self.caminho_salvamento)
            self.entry_dir.configure(state="disabled")

    def atualizar_status(self, texto):
        self.lbl_status.configure(text=texto)
        self.update_idletasks()

    def iniciar_busca(self):
        palavras_selecionadas = [termo.lower() for termo, var in self.check_vars.items() if var.get()]

        customizadas_raw = self.entry_custom.get()
        if customizadas_raw.strip():
            palavras_customizadas = [p.strip().lower() for p in customizadas_raw.split(",") if p.strip()]
            palavras_selecionadas.extend(palavras_customizadas)

        if not palavras_selecionadas:
            messagebox.showwarning("Aviso", "Selecione ou digite ao menos uma palavra-chave.")
            return

        if not os.path.exists(self.caminho_salvamento):
            try:
                os.makedirs(self.caminho_salvamento)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar a pasta de destino:\n{e}")
                return

        self.btn_iniciar.configure(state="disabled", text="BUSCANDO DADOS... (PODE DEMORAR)")
        self.btn_dir.configure(state="disabled")

        baixar_anexos = self.switch_baixar.get() == 1

        threading.Thread(target=self.processo_extracao,
                         args=(palavras_selecionadas, self.caminho_salvamento, baixar_anexos), daemon=True).start()

    def processo_extracao(self, palavras_chave, pasta_raiz, baixar_anexos):
        try:
            df_resultados = buscar_editais_pncp(palavras_chave, pasta_raiz, app_gui=self, baixar_arquivos=baixar_anexos)

            self.atualizar_status("Salvando CSV e finalizando...")
            total_salvos = salvar_em_csv(df_resultados, pasta_raiz)

            self.atualizar_status(f"Concluído! {total_salvos} editais capturados.")
            messagebox.showinfo("Sucesso",
                                f"Monitoramento finalizado.\n{total_salvos} editais ativos encontrados e salvos em:\n{pasta_raiz}")
        except Exception as e:
            self.atualizar_status("Erro durante a execução.")
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro no processo:\n{e}")
        finally:
            self.btn_iniciar.configure(state="normal", text="INICIAR MONITORAMENTO")
            self.btn_dir.configure(state="normal")


if __name__ == "__main__":
    app = AppMonitoramento()
    app.mainloop()