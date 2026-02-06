import pandas as pd
import numpy as np
from datetime import timedelta
import os
import plotly.io as pio
import plotly.express as px
import time


class Gerar_relatorio_permanencia:
    def __init__(self, diretorio_base):
        # Usamos o diretorio_base que vem do main.py para localizar os arquivos
        self.diretorio_base = diretorio_base



    def processar_dados(self, callback_status=None):

        def status(etapa, detalhe=""):
            if callback_status:
                callback_status(etapa, detalhe)
        try:
            # ETAPA 1: Carregamento
            status("Carregando arquivos...", "Lendo logs de permanência")
            path_bips = os.path.join(self.diretorio_base, "merge.xlsx")
            path_inscritos = os.path.join(self.diretorio_base,  "Disp WE 2026 - enriquecido.xlsx")

            df_logs = pd.read_excel(path_bips)
            df_disp = pd.read_excel(path_inscritos)


            # ETAPA 2: Processamento dos dados
            status("Processando dados...", "Calculando tempo de permanência")


            df_logs['Horario'] = pd.to_datetime(df_logs['Horario'])
            df_logs = df_logs.sort_values(['CD_PESSOA', 'Espaço', 'Horario'])
            df_logs['CD_PESSOA'] = df_logs['CD_PESSOA'].astype(str).str.extract(r'(\d+)').astype(int)


            sessions = []
            inconsistencias = []

            status("Analisando logs...", "Identificando sessões e inconsistências")
            for (pessoa, espaco), grupo in df_logs.groupby(['CD_PESSOA', 'Espaço']):
                estado = 'fora'
                entrada_atual = None

                for _, row in grupo.iterrows():
                    if row['Tipo_Bip'] == 'ENTRADA':
                        if estado == 'fora':
                            estado = 'dentro'
                            entrada_atual = row['Horario']
                        else:
                            inconsistencias.append({**row, 'Problema': 'ENTRADA duplicada'})

                    elif row['Tipo_Bip'] == 'SAÍDA':
                        if estado == 'dentro':
                            saida = row['Horario']
                            duracao = int((saida - entrada_atual).total_seconds() / 60)
                            sessions.append({
                                'CD_PESSOA': pessoa,
                                'Espaço': espaco,
                                'Entrada': entrada_atual,
                                'Saída': saida,
                                'Tempo_Minutos': duracao
                            })
                            estado = 'fora'
                            entrada_atual = None
                        else:
                            inconsistencias.append({**row, 'Problema': 'SAÍDA sem ENTRADA'})

                if estado == 'dentro':
                    inconsistencias.append({
                        'CD_PESSOA': pessoa,
                        'Espaço': espaco,
                        'Horario': entrada_atual,
                        'Tipo_Bip': 'ENTRADA',
                        'Problema': 'ENTRADA sem SAÍDA'
                    })



            df_sessions = pd.DataFrame(sessions)
            df_inconsistencias = pd.DataFrame(inconsistencias)

            
            df_sessions["Dia"] = np.where(df_sessions["Entrada"].dt.day == 4, "Dia 1", "Dia 2")
            df_resumido_por_espaco = df_sessions.groupby(["CD_PESSOA", "Espaço", "Dia"])["Tempo_Minutos"].sum().reset_index().sort_values('Tempo_Minutos', ascending=False)
            df_resumido_geral = df_sessions.groupby(["CD_PESSOA", "Dia"])["Tempo_Minutos"].sum().reset_index().sort_values('Tempo_Minutos', ascending=False)


            colunas = ['CD_PESSOA', 'Nome','Município' ,'Território' ,'Regional']
            df_disp = df_disp[colunas].drop_duplicates()

            status("Cruzando dados...", "Adicionando informações dos inscritos aos relatórios")
            df_inconsistencias = df_inconsistencias.merge(df_disp, on='CD_PESSOA', how='left')
            df_resumido_por_espaco = df_resumido_por_espaco.merge(df_disp, on='CD_PESSOA', how='left')
            df_resumido_geral = df_resumido_geral.merge(df_disp, on='CD_PESSOA', how='left')


            # Mudar ordem das colunas do df_resumido_por_espaco
            colunas_ordem_espaco = ['CD_PESSOA', 'Espaço', 'Dia', 'Nome', 'Município','Território', 'Regional', 'Tempo_Minutos']
            colunas_ordem_geral = ['CD_PESSOA', 'Dia', 'Nome', 'Município','Território', 'Regional', 'Tempo_Minutos']

            df_resumido_por_espaco = df_resumido_por_espaco[colunas_ordem_espaco]
            df_resumido_geral = df_resumido_geral[colunas_ordem_geral]

            caminho_saida_sessions = os.path.join(self.diretorio_base, "data/log_minutos.xlsx")
            caminho_saida_inconsistencias = os.path.join(self.diretorio_base, "data/inconsistencias.xlsx")
            caminho_saida_resumo_geral = os.path.join(self.diretorio_base, "data/resumo_geral.xlsx")
            caminho_saida_resumo_por_espaco = os.path.join(self.diretorio_base, "data/resumo_por_espaco.xlsx")

            df_sessions.to_excel(caminho_saida_sessions, index=False)
            df_inconsistencias.to_excel(caminho_saida_inconsistencias, index=False)
            df_resumido_geral.to_excel(caminho_saida_resumo_geral, index=False)
            df_resumido_por_espaco.to_excel(caminho_saida_resumo_por_espaco, index=False)




            media_por_territorio = df_resumido_geral.groupby('Território')['Tempo_Minutos'].mean().round(0).reset_index().sort_values('Tempo_Minutos', ascending=False)
            media_por_regional = df_resumido_geral.groupby('Regional')['Tempo_Minutos'].mean().round(0).reset_index().sort_values('Tempo_Minutos', ascending=False)


            max_território = media_por_territorio.loc[media_por_territorio['Tempo_Minutos'].idxmax()]

            sub_title = f"O território <b>{max_território['Território']}</b> possui a maior média de permanência no evento WE: {int(max_território['Tempo_Minutos'])} minutos."

            fig_territorio = px.bar(
                media_por_territorio, 
                text="Tempo_Minutos",
                x='Território', 
                y='Tempo_Minutos',
                title='Média de Permanência por Território',
                subtitle=sub_title,
                labels={
                    "Tempo_Minutos": "Média de Permanência (minutos)",
                    "Território": "Território"
                }
                )
                

            fig_regional = px.bar(
                media_por_regional, 
                text="Tempo_Minutos",
                x='Regional', 
                y='Tempo_Minutos',
                title='Média de Permanência por Regional',
                labels={
                    "Tempo_Minutos": "Média de Permanência (minutos)",
                    "Regional": "Regional"
                }
                )




            fig_box_regional = px.box(df_resumido_geral, x="Regional", y="Tempo_Minutos", title="Distribuição do Tempo em Minutos por Regional", labels={
                "Tempo_Minutos": "Tempo em Minutos",
                "Regional": "Regional"
                }   )        

            status("Gerando gráficos...", "Criando visualizações com Plotly")
            time.sleep(2)  # Simula o tempo para gerar os gráficos
            ## Salvar os gráficos em um arquivo HTML
            with open(os.path.join(self.diretorio_base, "data/Relatório Gráfico de Permanência no WE.html"), "w") as f:
                f.write(pio.to_html(fig_territorio, full_html=False, include_plotlyjs='cdn'))
                f.write(pio.to_html(fig_regional, full_html=False, include_plotlyjs='cdn'))
                f.write(pio.to_html(fig_box_regional, full_html=False, include_plotlyjs='cdn'))




            # # Exemplo de processamento: calcular tempo total por pessoa
            # df_logs['Tempo_Minutos'] = df_logs['Tempo_Saida'] - df_logs['Tempo_Entrada']
            # df_resumido = df_logs.groupby('CD_PESSOA')['Tempo_Minutos'].sum().reset_index()

            # # ETAPA 3: Salvamento do relatório
            # status("Salvando relatório...", "Gerando arquivo de saída")
            # path_saida = os.path.join(self.diretorio_base, "Relatório_Permanencia.csv")
            # df_resumido.to_csv(path_saida, index=False)

            # # ETAPA 4: Finalização
            # status("Processamento concluído!", f"Relatório salvo em: {path_saida}")

            return True, "Processamento concluído com sucesso!"



        except Exception as e:
            status("Erro durante o processamento", str(e))
            return False, f"Erro: {str(e)}"