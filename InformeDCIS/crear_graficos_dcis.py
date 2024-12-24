import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import numpy as np
import matplotlib.font_manager as fm

fe = fm.FontEntry(
    fname="/home/daniel/PycharmProjects/Fiscalia/InformeViolaciones/Metropolis-Regular.ttf",
    name="Metropolis"
)
fm.fontManager.ttflist.insert(0, fe)

plt.figure()
plt.rcParams.update({'font.size': 12,
                     'font.weight': 'bold',
                     'font.family': 'Metropolis'})

path = "/home/daniel/PycharmProjects/Fiscalia/InformeDCIS/"
file = "/home/daniel/PycharmProjects/Fiscalia/InformeDCIS/INFORMACIÓN CIS al 12 Diciembre 2024.xlsx"
datos = pd.read_excel(file, skiprows=1)
datos["AÑO"] = datos.iloc[:, 0].ffill().astype(int).astype(str)

ult_12_meses = datos.iloc[-16:-2]
ult_12_meses = ult_12_meses.rename(
    columns={"VINCULACIONES A PROCESO POR CUMPLIMIENTO DE OA": "Vinculaciones por OA",
             "VINCULACIONES A PROCESO CON DETENIDO": "Vinculaciones en flagrancia",
             "VINCULACIONES A PROCESO SIN DETENIDO": "Vinculaciones sin detenido",
             "OA CONCEDIDAS": "Órdenes concedidas",
             "OA CUMPLIMENTADAS": "Órdenes cumplimentadas",
             "OA SOLICITADAS": "Órdenes solicitadas"
             })
ult_12_meses = ult_12_meses.fillna(0)

ult_12_meses["MES"] = ult_12_meses["MES"].apply(lambda x: x[0:3])
ult_12_meses["AÑO"] = ult_12_meses["AÑO"].apply(lambda x: x[-2:])
ult_12_meses["MES-AÑO"] = ult_12_meses["MES"] + "-" + ult_12_meses["AÑO"]
ult_12_meses["Inicios Totales CD"] = (ult_12_meses["INICIOS C/D"] + ult_12_meses["RADICACIONES C/D"]).astype("int64")
ult_12_meses["Inicios Totales SD"] = (ult_12_meses["INICIOS S/D"] + ult_12_meses["RADICACIONES S/D"]).astype("int64")
ult_12_meses["Vinculaciones por investigación"] = (ult_12_meses["Vinculaciones por OA"] + ult_12_meses["Vinculaciones sin detenido"]).astype("int64")
ult_12_meses["Productividad CD"] = round((ult_12_meses["Vinculaciones en flagrancia"] / ult_12_meses["Inicios Totales CD"]).astype("float64") * 100, 1)
ult_12_meses["Productividad CD"] = round((ult_12_meses["Vinculaciones en flagrancia"] / ult_12_meses["Inicios Totales CD"]).astype("float64") * 100, 1)
ult_12_meses["Productividad SD"] = round((ult_12_meses["Vinculaciones por investigación"] / ult_12_meses["Inicios Totales SD"]).astype("float64") * 100, 1)
ult_12_meses["Productividad de solicitudes"] = round((ult_12_meses["Órdenes concedidas"] / ult_12_meses["Órdenes solicitadas"]).astype("float64") * 100, 1)
ult_12_meses["Productividad de cumplimientos"] = round((ult_12_meses["Órdenes cumplimentadas"] / ult_12_meses["Órdenes concedidas"]).astype("float64") * 100, 1)
ult_12_meses.replace([np.inf, -np.inf], np.nan, inplace=True)
# ult_12_meses.fillna(0, inplace=True)
ult_12_meses["Productividad CD Porcentaje"] = ult_12_meses["Productividad CD"].apply(lambda x: str(x) + "%")
ult_12_meses["Productividad SD Porcentaje"] = ult_12_meses["Productividad SD"].apply(lambda x: str(x) + "%")
ult_12_meses["Productividad de solicitudes Porcentaje"] = ult_12_meses["Productividad de solicitudes"].apply(lambda x: str(x) + "%")
ult_12_meses["Productividad de cumplimientos Porcentaje"] = ult_12_meses["Productividad de cumplimientos"].apply(lambda x: str(x) + "%")

ult_12_meses["Vinculaciones en flagrancia"] = ult_12_meses["Vinculaciones en flagrancia"].apply(lambda x: int(x))
ult_12_meses["Órdenes cumplimentadas"] = ult_12_meses["Órdenes cumplimentadas"].apply(lambda x: int(x))
ult_12_meses["Órdenes solicitadas"] = ult_12_meses["Órdenes solicitadas"].apply(lambda x: int(x))
ult_12_meses["Órdenes concedidas"] = ult_12_meses["Órdenes concedidas"].apply(lambda x: int(x))
df_prod_cd = ult_12_meses[["Inicios Totales CD", "Vinculaciones en flagrancia", "Productividad CD", "Productividad CD Porcentaje"]]
df_prod_cd["indice"] = range(1, 15)
df_prod_sd = ult_12_meses[["Inicios Totales SD", "Vinculaciones por investigación", "Productividad SD", "Productividad SD Porcentaje"]]
df_prod_sd["indice"] = range(1, 15)
df_prod_OA_sol = ult_12_meses[["Órdenes solicitadas", "Órdenes concedidas", "Productividad de solicitudes", "Productividad de solicitudes Porcentaje"]]
df_prod_OA_sol["indice"] = range(1, 15)
df_prod_OA_cum = ult_12_meses[["Órdenes concedidas", "Órdenes cumplimentadas", "Productividad de cumplimientos", "Productividad de cumplimientos Porcentaje"]]
df_prod_OA_cum["indice"] = range(1, 15)

# ax = ult_12_meses[["Inicios Totales CD","Vinculaciones en flagrancia" ]].plot(kind="bar", figsize=(12, 7), title="Violaciones con detenido. Últimos 12 meses.")
# ax.set_ylabel("Conteo")
# ax.set_xlabel("Mes")
# ax.set_xticklabels(ult_12_meses["MES"])
#
# ax2 = ax.twinx()
# ax2.set_ylabel("Porcentaje")
# ax2.plot(ax.get_xticks(), ult_12_meses["Productividad CD"], marker="o", color="green")
# ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
#
# # df = DataFrame(rand(3,2), columns=['A', 'B'])
# # ax = df.plot(table=True, kind='bar', title='Random')
# # for i, each in enumerate(ult_12_meses.index):
# #     for col in ult_12_meses.columns:
# #         y = ult_12_meses[each][col]
# #         ax.text(i, y, y)
#
# # for idx, row in ult_12_meses.iterrows():
# #     ax.annotate(row['Inicios Totales CD'], (idx, row['Inicios Totales CD']))
# # ax.text(ult_12_meses["Inicios Totales CD"], ult_12_meses[""], ult_12_meses["Inicios Totales CD"])

# Violaciones con detenido
ax = df_prod_cd.plot(kind="bar",
                     y=["Inicios Totales CD", "Vinculaciones en flagrancia"],
                     x="indice",
                     figsize=(14, 6),
                     color=["yellowgreen", "cornflowerblue"])
plt.title("""DCIS con detenido
Últimos 12 meses""", fontsize = 25)
ax.legend(fancybox=True, framealpha=0.5, bbox_to_anchor=(1, -0.15))
ax2 = ax.twinx()
ax2.set_ylabel("Porcentaje")
ax2.plot(ax.get_xticks(), df_prod_cd["Productividad CD"], marker="o", color="orange")
ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
ax2.set_ylim(0, 150)
ax2.legend(["Productividad CD"], fancybox=True, framealpha=0.5, bbox_to_anchor=(0.3, -0.15))
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax2.spines['top'].set_visible(False)
ax2.spines['right'].set_visible(False)
for idx, row in df_prod_cd.iterrows():
    ax.text(row["indice"] - 1.5, row['Inicios Totales CD'], row['Inicios Totales CD'])
    ax.text(row["indice"] - 0.75, row['Vinculaciones en flagrancia'], row['Vinculaciones en flagrancia'])
    ax2.text(row["indice"] - 1.25, row['Productividad CD'], row['Productividad CD Porcentaje'])
ax.set_ylabel("Conteo")
ax.set_xlabel("Mes")
ax.set_xticklabels(ult_12_meses["MES-AÑO"])
plt.savefig(f'{path}ProdCD.png', bbox_inches='tight', transparent=True)

# Violaciones sin detenido
ax = df_prod_sd.plot(kind="bar",
                     y=["Inicios Totales SD", "Vinculaciones por investigación"],
                     x="indice",
                     figsize=(14, 6),
                     color=["yellowgreen", "cornflowerblue"])
plt.title("""DCIS sin detenido
Últimos 12 meses""", fontsize = 25)
ax.legend(fancybox=True, framealpha=0.5, bbox_to_anchor=(1, -0.15))
ax2 = ax.twinx()
ax2.set_ylabel("Porcentaje")
ax2.plot(ax.get_xticks(), df_prod_sd["Productividad SD"], marker="o", color="orange")
ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
ax2.set_ylim(0, 150)
ax2.legend(["Productividad SD"], fancybox=True, framealpha=0.5, bbox_to_anchor=(0.3, -0.15))
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax2.spines['top'].set_visible(False)
ax2.spines['right'].set_visible(False)
for idx, row in df_prod_sd.iterrows():
    ax.text(row["indice"] - 1.5, row['Inicios Totales SD'], row['Inicios Totales SD'])
    ax.text(row["indice"] - 0.75, row['Vinculaciones por investigación'], row['Vinculaciones por investigación'])
    ax2.text(row["indice"] - 1.25, row['Productividad SD'], row['Productividad SD Porcentaje'])
ax.set_ylabel("Conteo")
ax.set_xlabel("Mes")
ax.set_xticklabels(ult_12_meses["MES-AÑO"])
plt.savefig(f'{path}ProdSD.png', bbox_inches='tight', transparent=True)

# OA solicitadas
ax = df_prod_OA_sol.plot(kind="bar",
                     y=["Órdenes solicitadas", "Órdenes concedidas"],
                     x="indice",
                         figsize=(14, 6),
                         color=["yellowgreen", "cornflowerblue"])
plt.title("""Productividad de solicitudes
Últimos 12 meses""", fontsize = 25)
ax.legend(fancybox=True, framealpha=0.5, bbox_to_anchor=(1, -0.15))
ax2 = ax.twinx()
ax2.set_ylabel("Porcentaje")
ax2.plot(ax.get_xticks(), df_prod_OA_sol["Productividad de solicitudes"], marker="o", color="orange")
ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
ax2.set_ylim(0, 150)
ax2.legend(["Productividad de solicitudes"], fancybox=True, framealpha=0.5, bbox_to_anchor=(0.3, -0.15))
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax2.spines['top'].set_visible(False)
ax2.spines['right'].set_visible(False)
for idx, row in df_prod_OA_sol.iterrows():
    ax.text(row["indice"] - 1.5, row['Órdenes solicitadas'], row['Órdenes solicitadas'])
    ax.text(row["indice"] - 0.75, row['Órdenes concedidas'], row['Órdenes concedidas'])
    ax2.text(row["indice"] - 1.25, row['Productividad de solicitudes'], row['Productividad de solicitudes Porcentaje'])
ax.set_ylabel("Conteo")
ax.set_xlabel("Mes")
ax.set_xticklabels(ult_12_meses["MES-AÑO"])
plt.savefig(f'{path}ProdOASol.png', bbox_inches='tight', transparent=True)

# OA cumplimentadas
ax = df_prod_OA_cum.plot(kind="bar",
                         y=["Órdenes concedidas", "Órdenes cumplimentadas"],
                         x="indice",
                         figsize=(14, 6),
                         color=["yellowgreen", "cornflowerblue"])
plt.title("""Productividad de cumplimientos
Últimos 12 meses""", fontsize = 25)
ax.legend(fancybox=True, framealpha=0.5, bbox_to_anchor=(1, -0.15))
ax2 = ax.twinx()
ax2.set_ylabel("Porcentaje")
ax2.plot(ax.get_xticks(), df_prod_OA_cum["Productividad de cumplimientos"], marker="o", color="orange")
ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
ax2.set_ylim(0, 150)
ax2.legend(["Productividad de cumplimientos"], fancybox=True, framealpha=0.5, bbox_to_anchor=(0.3, -0.15))
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax2.spines['top'].set_visible(False)
ax2.spines['right'].set_visible(False)
for idx, row in df_prod_OA_cum.iterrows():
    ax.text(row["indice"] - 1.5, row['Órdenes concedidas'], row['Órdenes concedidas'])
    ax.text(row["indice"] - 0.75, row['Órdenes cumplimentadas'], row['Órdenes cumplimentadas'])
    ax2.text(row["indice"] - 1.25, row['Productividad de cumplimientos'], row['Productividad de cumplimientos Porcentaje'])
ax.set_ylabel("Conteo")
ax.set_xlabel("Mes")
ax.set_xticklabels(ult_12_meses["MES-AÑO"])
plt.savefig(f'{path}ProdOACum.png', bbox_inches='tight', transparent=True)
