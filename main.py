import os
import csv
import datetime
import sys
from pathlib import Path
import customtkinter as ctk
from tkinter import messagebox
from PIL import Image
import pandas as pd
from tkinter import filedialog
import threading
import time
import sys
import os
from tratamento_bi import ProcessadorBI
from tratamento_logs_evento import Gerar_relatorio_permanencia

# Configuração de temas do CustomTkinter
ctk.set_appearance_mode("dark")  # Opções: "dark", "light"
ctk.set_default_color_theme("blue")

def obter_caminho(caminho_relativo):
    """ Função utilitária para localizar recursos (imagens, dados) """
    try:
        # Caminho quando compilado pelo PyInstaller
        base_path = sys._MEIPASS
    except Exception:
        # Caminho quando rodando como script .py
        base_path = os.path.abspath(".")
    return os.path.join(base_path, caminho_relativo)



class AppCredenciamento(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuração da Janela
        self.title("Sistema de Credenciamento")
        self.geometry("900x750")
        try:
            self.iconbitmap(obter_caminho("images/p.ico"))
        except Exception:
            pass 

        
        # Gerenciamento de Caminhos (compatível com .exe)
        if getattr(sys, 'frozen', False):
            self.diretorio_base = Path(sys.executable).parent
        else:
            self.diretorio_base = Path(__file__).parent       
 
        self.pasta_data = self.diretorio_base / "data"
        self.pasta_data.mkdir(parents=True, exist_ok=True)
        
        data_atual = datetime.datetime.now().strftime("%Y-%m-%d")
        self.caminho_csv = self.pasta_data / f"credenciamento_{data_atual}.csv"
        self.criar_arquivo_se_nao_existir()

        # Criando as imagens diretamente com a função
        self.meu_icone_merge = ctk.CTkImage(
            light_image=Image.open(obter_caminho("images/merge.png")),
            dark_image=Image.open(obter_caminho("images/merge.png")), 
            size=(20, 20)
        )  
        
        self.meu_icone_trash = ctk.CTkImage(
            light_image=Image.open(obter_caminho("images/trash.png")),
            dark_image=Image.open(obter_caminho("images/trash.png")), 
            size=(20, 20)
        )  
        
        self.meu_icone_graph = ctk.CTkImage(
            light_image=Image.open(obter_caminho("images/graph.png")),
            dark_image=Image.open(obter_caminho("images/graph.png")), 
            size=(20, 20)
        )

        self.meu_icone_relogio = ctk.CTkImage(
            light_image=Image.open(obter_caminho("images/relogio.png")),
            dark_image=Image.open(obter_caminho("images/relogio.png")),
            size=(20, 20)
        )

        # Para a logo (se for usar em um Label depois)
        self.logo_sebrae_img = Image.open(obter_caminho("images/we_logo.png"))


        # --- INTERFACE GRÁFICA (UI) ---
        
        # Grid Layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # # Botão de Configurações
        # self.btn_config = ctk.CTkButton(self, text="", image=self.meu_icone_config, compound="left", command=self.abrir_configuracoes, width=20, height=20)
        # self.btn_config.grid(row=0, column=0, pady=5, sticky="e", padx=10)
           


        self.switch_tipo_bip = ctk.CTkSwitch(self, text="Modo Entrada", font=ctk.CTkFont(weight="bold", family="Arial")) # , command=self.atualizar_texto_switch
        self.switch_tipo_bip.select()
        self.switch_tipo_bip.grid(row=0, column=0, pady=5, sticky="w", padx=10)

        self.combobox_dia = ctk.CTkComboBox(self, values=["", "Dia 1", "Dia 2"], width=130, state="readonly")
        self.combobox_dia.grid(row=0, column=0, pady=1, padx=150, sticky="w")
        self.combobox_dia.set("Selecione o Dia") 

        self.combobox_espaco = ctk.CTkComboBox(self, values=["", "Auditório", "Sala 1 a 6", "Sala 7 a 10"], width=150, state="readonly")
        self.combobox_espaco.grid(row=0, column=0, pady=1, padx=300, sticky="w")
        self.combobox_espaco.set("Selecione o Espaço") 



        ## Adicionar logo WE na interface
        self.logo_sebrae = ctk.CTkImage(light_image=Image.open(obter_caminho("images/we_logo.png")),dark_image=Image.open(obter_caminho("images/we_logo.png")), size=(200, 130))  
        self.lbl_logo = ctk.CTkLabel(self, image=self.logo_sebrae, text="")
        self.lbl_logo.grid(row=1, column=0,  pady=(20, 5))

        # Título e Status
        # self.lbl_titulo = ctk.CTkLabel(self, text="WE 2026 - SEBRAE PR", font=ctk.CTkFont(size=24, weight="bold"))
        # self.lbl_titulo.grid(row=1, column=0, pady=(20, 5))

        self.lbl_instrucao = ctk.CTkLabel(self, text="Posicione o cursor no campo abaixo para coletar os QR Codes", font=ctk.CTkFont(size=15))
        self.lbl_instrucao.grid(row=2, column=0, pady=5)
        
        self.lbl_developer = ctk.CTkLabel(self, text="Desenvolvido por Patriky E. Galvão Mirkoski", font=ctk.CTkFont(size=10))
        self.lbl_developer.grid(row=6, column=0, pady=5)

        ## Incluir a versão do sistema
        self.lbl_versao = ctk.CTkLabel(self, text="Versão 2.2", font=ctk.CTkFont(size=10))
        self.lbl_versao.grid(row=6, column=0, pady=5, sticky="e", padx=20)

        # Campo de Entrada Principal (Onde os coletores vão "digitar")
        self.entry_codigo = ctk.CTkEntry(self, placeholder_text="Iniciar leitura...", text_color="#39FF14", 
                                        width=600, height=130, font=ctk.CTkFont(size=60, family="Arial", weight="bold"), corner_radius=20, 
                                        justify="center")

        self.entry_codigo.grid(row=3, column=0, pady=10)
        self.entry_codigo.bind("<Return>", self.processar_leitura) # Dispara ao apertar Enter (enviado pelo coletor)
        self.entry_codigo.focus_set()

        # # Área de Log/Histórico
        self.textbox_log = ctk.CTkTextbox(self, width=600, height=200, font=ctk.CTkFont(family="Consolas", size=14))
        self.textbox_log.grid(row=4, column=0, padx=20, pady=10, sticky="nsew")
        self.textbox_log.configure(state="disabled")

        # # Frame Inferior (Status e Total)
        self.status_frame = ctk.CTkFrame(self, height=40)
        self.status_frame.grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 20))
        

        self.lbl_total = ctk.CTkLabel(self.status_frame, text="Total de Registros: 0", font=ctk.CTkFont(weight="bold"))
        self.lbl_total.pack(side="left", padx=20, pady=5)



        self.btn_preparar_bi = ctk.CTkButton(self.status_frame, text="Relatório de Tempo", image=self.meu_icone_relogio, compound="left", command=self.calcular_media_permanencia, fg_color="#098bd6", hover_color="#505050", font=ctk.CTkFont(weight="bold"))
        self.btn_preparar_bi.pack(side="right", padx=10, pady=5)

        self.btn_preparar_bi = ctk.CTkButton(self.status_frame, text="Preparar BI", image=self.meu_icone_graph, compound="left", command=self.iniciar_processamento_bi, fg_color="#098bd6", hover_color="#505050", font=ctk.CTkFont(weight="bold"))
        self.btn_preparar_bi.pack(side="right", padx=10, pady=5 )

        # Botão de Configuração/Agrupar no canto ou em um Frame de ferramentas
        self.btn_agrupar = ctk.CTkButton(self.status_frame,  text="Agrupar CSVs", image=self.meu_icone_merge, compound="left", command=self.agrupar_csv, fg_color="#098bd6", hover_color="#505050", font=ctk.CTkFont(weight="bold"))
        self.btn_agrupar.pack(side="right", padx=10, pady=5) # 'ne' coloca no canto superior direito     

        ## Botão para limpar os registros do dia atual
        self.btn_limpar = ctk.CTkButton(self.status_frame, text="Limpar registros", image=self.meu_icone_trash, compound="left", command=self.limpar_registros, fg_color="#d60b0b", hover_color="#505050", font=ctk.CTkFont(weight="bold"))
        self.btn_limpar.pack(side="right", padx=10, pady=5) 


        self.protocol("WM_DELETE_WINDOW", self.sair)
        

        # Carregar total inicial
        self.atualizar_contador_total()



    # --- LÓGICA DE NEGÓCIO ---

    def obter_caminho(caminho_relativo):
        """ Retorna o caminho absoluto para recursos, compatível com dev e PyInstaller """
        try:
            # O PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, caminho_relativo)

    def criar_arquivo_se_nao_existir(self):
        if not self.caminho_csv.exists():
            with open(self.caminho_csv, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(["CD_PESSOA", "Horario", "Dia", "Tipo_Bip", "Espaço"])  # Cabeçalho

    def processar_leitura(self, event=None):
        codigo = self.entry_codigo.get().strip()
        
        if not codigo:
            return


        horario = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        dia = self.combobox_dia.get()
        espaco = self.combobox_espaco.get()
        tipo_bip = self.switch_tipo_bip.get()
        if tipo_bip == 1:
            tipo_bip_texto = "ENTRADA"
        else:
            tipo_bip_texto = "SAÍDA"

        try:
            # Escrita no CSV com tratamento de erro
            with open(self.caminho_csv, 'a', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow([codigo, horario, dia, tipo_bip_texto, espaco])
            
            self.adicionar_log(f" {horario} - Código {codigo}")
            self.atualizar_contador_total()
            
        except PermissionError:
            messagebox.showerror("Erro Crítico", "O arquivo CSV está aberto em outro programa (Excel?). Feche-o para continuar!")
        except Exception as e:
            self.adicionar_log(f"❌ Erro ao salvar: {str(e)}")
        
        # Limpa campo e garante foco para o próximo coletor
        self.entry_codigo.delete(0, 'end')
        self.entry_codigo.focus_set()

    def adicionar_log(self, mensagem):
        self.textbox_log.configure(state="normal")
        self.textbox_log.insert("0.0", mensagem + "\n") # Insere no topo
        self.textbox_log.configure(state="disabled")

    def atualizar_contador_total(self):
        try:
            with open(self.caminho_csv, 'r', encoding='utf-8') as f:

                linhas = sum(1 for line in f) - 1 # Desconta cabeçalho
                self.lbl_total.configure(text=f"Total de Registros: {max(0, linhas)}")
        except:
            print("Erro ao ler o arquivo CSV para contagem.")

    def limpar_registros(self):

        ## Verificar se o usuário realmente quer limpar o log visual
        resposta = messagebox.askyesno("Confirmação", f"Tem certeza que deseja limpar os registros do arquivo {self.caminho_csv.name}?")
        if resposta:
            try:
                with open(self.caminho_csv, 'w', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(["CD_PESSOA", "Horario", "Dia", "Tipo_Bip", "Espaço"])
            except:
                print("Erro ao ler o arquivo CSV para contagem.")        


            self.textbox_log.configure(state="normal") 
            self.textbox_log.delete("0.0", "end") 
            self.textbox_log.configure(state="disabled")
            self.atualizar_contador_total()

    def abrir_configuracoes(self):
        messagebox.showinfo("Configurações", "Aqui você pode adicionar configurações futuras.")

    def sair(self):

        resposta = messagebox.askyesno("Confirmação", "Tem certeza que deseja sair do aplicativo?")
        if resposta:
            self.destroy()

    def help_agrupar(self):
        """Exibe as instruções em uma caixa de diálogo moderna."""
        instrucoes = (
            "1. Selecione dois ou mais arquivos CSV.\n"
            "2. Certifique-se de que as colunas sejam iguais.\n"
            "3. Escolha o local e o formato (.xlsx ou .csv) para salvar.\n\n"
            "O sistema irá unir todos os dados em um único arquivo."
        )
        messagebox.showinfo("Ajuda: Agrupar Arquivos", instrucoes)

    def agrupar_csv(self):
        """Lógica de agrupamento integrada à interface."""
        # 1. Mostrar ajuda antes de começar
        self.help_agrupar()

        # 2. Seleção de arquivos (filedialog funciona direto no CTk)
        file_paths = filedialog.askopenfilenames(
            title="Selecionar arquivos CSV para agrupar",
            filetypes=[("Arquivos CSV", "*.csv")]
        )
        
        if not file_paths:
            return # Usuário cancelou

        try:
            # 3. Processamento com Pandas
            df_list = []
            for path in file_paths:
                # Usamos utf-8 para manter compatibilidade com o que salvamos
                df = pd.read_csv(path, encoding='utf-8') 
                df_list.append(df)
            
            df_agrupado = pd.concat(df_list, ignore_index=True)

            # 4. Salvar arquivo
            caminho_saida = filedialog.asksaveasfilename(
                title="Salvar arquivo agrupado como",
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")],
                # initialfile=f"consolidado_{datetime.datetime.now().strftime('%Y%m%d')}"
                initialfile="merge"
            )

            if caminho_saida:
                if caminho_saida.endswith('.xlsx'):
                    df_agrupado.to_excel(caminho_saida, index=False)
                else:
                    df_agrupado.to_csv(caminho_saida, index=False, encoding='utf-8')
                
                messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{caminho_saida}")
                self.adicionar_log(f"📂 Arquivos agrupados em: {os.path.basename(caminho_saida)}")
                
        except Exception as e:
            messagebox.showerror("Erro no Agrupamento", f"Ocorreu um erro: {str(e)}")

    def iniciar_processamento_bi(self):
        """Cria a janela de status e inicia a thread de processamento."""
        # Criar Janela de Status (Pop-up)
        self.win_status = ctk.CTkToplevel(self)
        self.win_status.title("Processamento BI - 2026")
        self.win_status.geometry("400x250")
        self.win_status.attributes("-topmost", True) # Fica na frente
        self.win_status.grab_set() # Bloqueia a janela principal

        self.lbl_etapa = ctk.CTkLabel(self.win_status, text="Iniciando...", font=("Arial", 14, "bold"))
        self.lbl_etapa.pack(pady=20)

        # Barra de progresso (modo indeterminado para SQL demorado)
        self.progress = ctk.CTkProgressBar(self.win_status, width=300)
        self.progress.pack(pady=10)
        self.progress.configure(mode="indeterminado")
        self.progress.start()

        self.lbl_detalhes = ctk.CTkLabel(self.win_status, text="Isso pode levar alguns minutos.\nNão feche o programa.", font=("Arial", 11))
        self.lbl_detalhes.pack(pady=10)

        # Criamos uma função interna (wrapper) para rodar na Thread
        def tarefa_completa():
            processador = ProcessadorBI(self.diretorio_base)
            # 1. Roda o processamento pesado
            sucesso, mensagem = processador.processar_dados(self.atualizar_status_ui)
            
            # 2. Quando terminar, chama a finalização na Thread Principal (UI)
            self.after(0, lambda: self.encerrar_janela_processamento(sucesso, mensagem))

        # Inicia a Thread única
        thread = threading.Thread(target=tarefa_completa)
        thread.start()

    def calcular_media_permanencia(self):

        """Cria a janela de status e inicia a thread de processamento."""
        # Criar Janela de Status (Pop-up)
        self.win_status = ctk.CTkToplevel(self)
        self.win_status.title("Relatório de Permanência - 2026")
        self.win_status.geometry("400x250")
        self.win_status.attributes("-topmost", True) # Fica na frente
        self.win_status.grab_set() # Bloqueia a janela principal

        self.lbl_etapa = ctk.CTkLabel(self.win_status, text="Iniciando...", font=("Arial", 14, "bold"))
        self.lbl_etapa.pack(pady=20)

        # Barra de progresso (modo indeterminado para SQL demorado)
        self.progress = ctk.CTkProgressBar(self.win_status, width=300)
        self.progress.pack(pady=10)
        self.progress.configure(mode="indeterminado")
        self.progress.start()

        self.lbl_detalhes = ctk.CTkLabel(self.win_status, text="Isso pode levar alguns minutos.\nNão feche o programa.", font=("Arial", 11))
        self.lbl_detalhes.pack(pady=10)

        # Criamos uma função interna (wrapper) para rodar na Thread
        def iniciar_relatorio():
            relatorio = Gerar_relatorio_permanencia(self.diretorio_base)
            # # 1. Roda o processamento pesado
            sucesso, mensagem = relatorio.processar_dados(self.atualizar_status_ui)
            
            if sucesso:
                self.adicionar_log("Relatório de permanência gerado com sucesso.")
                self.after(0, lambda: self.encerrar_janela_processamento(True, "Relatórios de permanência gerados com sucesso!", relatorio_gerado=True))
            else:
                self.adicionar_log("Erro ao gerar o relatório de permanência.")
                self.after(0, lambda: self.encerrar_janela_processamento(False, "Falha ao gerar os relatórios de permanência."))

            # # 2. Quando terminar, chama a finalização na Thread Principal (UI)

            # 2. Quando terminar, chama a finalização na Thread Principal (UI)
            


        # Inicia a Thread única
        thread = threading.Thread(target=iniciar_relatorio)
        thread.start()



    def encerrar_janela_processamento(self, sucesso, mensagem, relatorio_gerado=False):
            """Fecha a janela e avisa o usuário."""
            if hasattr(self, 'win_status') and self.win_status.winfo_exists():
                self.win_status.destroy()
            
            if sucesso:
                messagebox.showinfo("Sucesso", F"{mensagem}")
                self.adicionar_log("BI Atualizado.")
                ## abrir o arquivo HTML gerado
                if relatorio_gerado:
                    caminho_html = os.path.join(self.diretorio_base, "data/Relatório Gráfico de Permanência no WE.html")
                    if os.path.exists(caminho_html):
                        os.startfile(caminho_html)

            else:
                messagebox.showerror("Erro", f"Falha: {mensagem}")

    def atualizar_status_ui(self, etapa, detalhe=""):
        """Função para atualizar a UI de dentro da Thread com segurança."""
        self.lbl_etapa.configure(text=etapa)
        if detalhe:
            self.adicionar_log(f"🔄 {etapa} - {detalhe}")
            self.lbl_detalhes.configure(text=detalhe)
        self.update_idletasks() # Força a interface a se redesenhar


if __name__ == "__main__":
    app = AppCredenciamento()
    app.mainloop()