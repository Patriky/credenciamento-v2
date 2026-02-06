import pandas as pd
import numpy as np
from datetime import timedelta
import os

class ProcessadorBI:
    def __init__(self, diretorio_base):
        # Usamos o diretorio_base que vem do main.py para localizar os arquivos
        self.diretorio_base = diretorio_base

    def processar_dados(self, callback_status=None):
        """
        Executa a lógica de tratamento de dados.
        callback_status: função para atualizar o texto na interface.
        """
        def status(etapa, detalhe=""):
            if callback_status:
                callback_status(etapa, detalhe)

        try:
            # ETAPA 1: Carregamento
            status("Carregando arquivos...", "Lendo bips e base de inscritos")
            
            # Caminhos relativos ao diretório do executável
            path_bips = os.path.join(self.diretorio_base, "merge.xlsx")
            path_inscritos = os.path.join(self.diretorio_base,  "Disp WE 2026 - enriquecido.xlsx")

            df_bips = pd.read_excel(path_bips)
            df_inscritos = pd.read_excel(path_inscritos)

            # ETAPA 2: Formatação de Horários
            status("Processando horários...", "Arredondando para intervalos de 5min")
            df_bips["horario_formatado"] = df_bips['Horario'].values.astype('<M8[m]')
            df_bips["horario_formatado"] = pd.to_datetime(df_bips["horario_formatado"]).dt.floor('5min')

            # ETAPA 3: Identificação de Presença
            status("Cruzando dados...", "Identificando pessoas presentes")
            array_pessoas_presentes = df_bips['CD_PESSOA'].unique()
            df_inscritos["PRESENTE_WE"] = np.where(df_inscritos['CD_PESSOA'].isin(array_pessoas_presentes), 1, 0)

            # ETAPA 4: Merge e Lógica de Dias
            status("Finalizando cálculos...", "Aplicando lógica de Dia 1 e Dia 2")
            df_bips_cli = df_inscritos.merge(df_bips, on='CD_PESSOA', how="left")

            df_bips_cli["Dia 1"] = np.where(df_bips_cli["Dia"] == "Dia 1", 1, 0)
            df_bips_cli["Dia 2"] = np.where(df_bips_cli["Dia"] == "Dia 2", 1, 0)

            
            # Truncar a data para o dia "2026-01-01", mantendo o horário
            df_bips_cli["horario_formatado_modificada"] = pd.to_datetime(
                "2026-01-01 " + df_bips_cli["horario_formatado"].dt.strftime("%H:%M:%S")
            )

            # ETAPA 5: Salvamento
            status("Salvando resultado...", "Gerando ARQUIVO_BASE_POWER_BI.xlsx")
            caminho_saida = os.path.join(self.diretorio_base, "ARQUIVO_BASE_POWER_BI.xlsx")
            df_bips_cli.to_excel(caminho_saida, index=False)

            return True, "Processamento concluído com sucesso!"

        except Exception as e:
            return False, f"Erro no processamento: {str(e)}"