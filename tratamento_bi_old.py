import pandas as pd
import numpy as np
from datetime import timedelta

df_bips = pd.read_excel("data/merge.xlsx")
df_inscritos = pd.read_excel("Disp WE 2026 - enriquecido.xlsx")


df_bips["horario_formatado"] = df_bips['Horario'].values.astype('<M8[m]')
df_bips["horario_formatado"] = df_bips["horario_formatado"].dt.floor('5min')


array_pessoas_presentes = df_bips['CD_PESSOA'].unique()

df_inscritos["PRESENTE_WE"] = np.where(df_inscritos['CD_PESSOA'].isin(array_pessoas_presentes), 1,0)


df_bips_cli = df_inscritos.merge(df_bips,  left_on='CD_PESSOA', right_on='CD_PESSOA',how = "left")

df_bips_cli["Dia 1"] = np.where(df_bips_cli["Dia"] == "Dia 1", 1,0)
df_bips_cli["Dia 2"] = np.where(df_bips_cli["Dia"] == "Dia 2", 1,0)

df_bips_cli["horario_formatado_modificada"] = np.where(df_bips_cli["Dia"] == "Dia 1", df_bips_cli["horario_formatado"] + timedelta(days=1), df_bips_cli["horario_formatado"])

df_bips_cli.to_excel("ARQUIVO_BASE_POWER_BI.xlsx", index=False)




