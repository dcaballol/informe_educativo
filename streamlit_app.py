
import streamlit as st
import pandas as pd
import os
import matplotlib.pyplot as plt
from io import BytesIO
import base64
from jinja2 import Environment, FileSystemLoader
from datetime import datetime

# === CONFIG ===
base_path = r"C:/Users/Daniel Contreras/Desktop/proyecto_reportes_educacion/data"
template_name = "template_reporte_mejorado.html"

# === CARGA DE DATOS ===
df_matricula = pd.read_excel(os.path.join(base_path, "matricula", "Matricula_EE_resumen.xlsx"))
df_asistencia = pd.read_excel(os.path.join(base_path, "asistencia", "Asistencia_EE_resumen.xlsx"))
df_simce = pd.read_excel(os.path.join(base_path, "simce", "SIMCE_puntajes_resumen.xlsx"))

df_matricula.rename(columns={"AGNO": "Año"}, inplace=True)
df_asistencia.rename(columns={"AGNO": "Año"}, inplace=True)
df_simce.rename(columns={"ANIO": "Año"}, inplace=True)

# === STREAMLIT ===
st.title("📘 Generador de Informes Educativos")

rbd_unicos = sorted(set(df_matricula["RBD"]) | set(df_asistencia["RBD"]) | set(df_simce["RBD"]))
rbd_seleccionados = st.multiselect("Selecciona RBD(s)", ["[TOTAL]"] + rbd_unicos)

tipos_info = st.multiselect("Selecciona tipos de información", ["Matrículas", "Asistencia", "SIMCE"], default=["Matrículas", "Asistencia", "SIMCE"])

años_m = sorted(df_matricula["Año"].unique())
años_a = sorted(df_asistencia["Año"].unique())
años_s = sorted(df_simce["Año"].unique())

años_m_select = st.multiselect("Años Matrícula", años_m, default=años_m) if "Matrículas" in tipos_info else []
años_a_select = st.multiselect("Años Asistencia", años_a, default=años_a) if "Asistencia" in tipos_info else []
años_s_select = st.multiselect("Años SIMCE", años_s, default=años_s) if "SIMCE" in tipos_info else []

# === FUNCIONES ===
def comparar_promedios(df, variable):
    df_sorted = df.sort_values("Año", ascending=False).copy()
    df_sorted["Año"] = df_sorted["Año"].astype(int)
    textos = []
    for i in range(1, len(df_sorted)):
        año_actual = df_sorted.iloc[0]["Año"]
        año_pasado = df_sorted.iloc[i]["Año"]
        val_actual = df_sorted.iloc[0][variable]
        val_pasado = df_sorted.iloc[i][variable]

        if val_pasado == 0:
            pct = 0
        else:
            pct = round(((val_actual - val_pasado) / val_pasado) * 100, 1)

        cambio = "aumentó" if val_actual > val_pasado else "disminuyó" if val_actual < val_pasado else "se mantuvo igual"
        textos.append(f"{variable} {cambio} {abs(pct)}% respecto a {año_pasado}")
    return " · ".join(textos)

def generar_grafico(df, x, y, title, color):
    df[x] = df[x].astype(int)
    fig, ax = plt.subplots()
    ax.plot(df[x], df[y], marker="o", color=color)
    for i, row in df.iterrows():
        val = row[y]
        txt = f"{val:.1f}%" if y == "Asistencia" else f"{val:,.0f}"
        ax.annotate(txt, (row[x], row[y]), textcoords="offset points", xytext=(0,10), ha="center", fontsize=8)
    ax.set_title(title)
    ax.set_xlabel("Año")
    ax.set_ylabel(y)
    ax.grid(True, linestyle="--", alpha=0.6)
    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()

def resumen_territorial():
    resumen = {}
    if not df_matricula.empty:
        t = df_matricula.groupby("Año")["Matricula"].sum().reset_index().sort_values("Año", ascending=False).head(2)
        t["Año"] = t["Año"].astype(int)
        t["Matricula"] = t["Matricula"].round(0).astype(int)
        resumen["matricula"] = t.to_html(index=False)
    if not df_asistencia.empty:
        t = df_asistencia.groupby("Año")["Asistencia"].mean().reset_index().sort_values("Año", ascending=False).head(2)
        t["Año"] = t["Año"].astype(int)
        t["Asistencia"] = (t["Asistencia"] * 100).round(1)
        resumen["asistencia"] = t.to_html(index=False)
    if not df_simce.empty:
        simce = {}
        for nivel in df_simce["NIVEL"].unique():
            df_n = df_simce[df_simce["NIVEL"] == nivel]
            t = df_n.groupby("Año")[["Lectura", "Matemática"]].mean().reset_index().sort_values("Año", ascending=False).head(2)
            t["Año"] = t["Año"].astype(int)
            t[["Lectura", "Matemática"]] = t[["Lectura", "Matemática"]].round(0).astype(int)
            t.insert(1, "NIVEL", nivel)
            simce[nivel] = t.to_html(index=False)
        resumen["simce"] = simce
    return resumen

# === GENERACIÓN ===
if st.button("🧾 Generar informe completo y descargar"):
    resumen_data = resumen_territorial()

    for rbd in rbd_seleccionados:
        datos, textos, graficos = {}, {}, {}
        textos["simce"] = {}
        rbd_label = str(rbd)
        hoy = datetime.today().strftime("%d/%m/%Y")

        if "Matrículas" in tipos_info:
            df = df_matricula if rbd == "[TOTAL]" else df_matricula[df_matricula["RBD"] == rbd]
            df = df[df["Año"].isin(años_m_select)].copy()
            df["Matricula"] = df["Matricula"].round(0).astype(int)
            if rbd == "[TOTAL]":
                tabla = df.groupby("Año")["Matricula"].sum().reset_index()
            else:
                tabla = df
            tabla["Año"] = tabla["Año"].astype(int)
            textos["matricula"] = comparar_promedios(tabla, "Matricula")
            datos["matricula"] = tabla.sort_values("Año", ascending=False).to_html(index=False)
            graficos["matricula"] = generar_grafico(tabla, "Año", "Matricula", "Evolución Matrícula", "#3498db")

        if "Asistencia" in tipos_info:
            df = df_asistencia if rbd == "[TOTAL]" else df_asistencia[df_asistencia["RBD"] == rbd]
            df = df[df["Año"].isin(años_a_select)].copy()
            df["Asistencia"] = (df["Asistencia"] * 100).round(1)
            tabla = df.groupby("Año")["Asistencia"].mean().reset_index() if rbd == "[TOTAL]" else df
            tabla["Año"] = tabla["Año"].astype(int)
            textos["asistencia"] = comparar_promedios(tabla, "Asistencia")
            asistencia_html = tabla.sort_values("Año", ascending=False).to_html(index=False)
            if rbd != "[TOTAL]":
                asistencia_html += '<div class="nota">Año 2025 asistencia último mes reportada.</div>'
            datos["asistencia"] = asistencia_html
            graficos["asistencia"] = generar_grafico(tabla, "Año", "Asistencia", "Evolución Asistencia (%)", "#e67e22")

        if "SIMCE" in tipos_info:
            df = df_simce if rbd == "[TOTAL]" else df_simce[df_simce["RBD"] == rbd]
            df = df[df["Año"].isin(años_s_select)]
            if not df.empty:
                tablas_simce, graficos_simce = {}, {}
                for nivel in df["NIVEL"].unique():
                    df_n = df[df["NIVEL"] == nivel].copy()
                    df_n = df_n[["Año", "Lectura", "Matemática"]].groupby("Año").mean().reset_index()
                    df_n["Año"] = df_n["Año"].astype(int)
                    df_n[["Lectura", "Matemática"]] = df_n[["Lectura", "Matemática"]].round(0).astype(int)
                    textos["simce"][nivel] = comparar_promedios(df_n, "Lectura") + " · " + comparar_promedios(df_n, "Matemática")
                    tablas_simce[nivel] = df_n.sort_values("Año", ascending=False).to_html(index=False)
                    fig, ax = plt.subplots()
                    ax.plot(df_n["Año"], df_n["Lectura"], marker="o", label="Lectura")
                    ax.plot(df_n["Año"], df_n["Matemática"], marker="s", label="Matemática")
                    ax.set_title(f"Evolución SIMCE - {nivel}")
                    ax.set_xlabel("Año")
                    ax.set_ylabel("Puntaje")
                    ax.legend()
                    ax.grid(True, linestyle="--", alpha=0.6)
                    fig.tight_layout()
                    buf = BytesIO()
                    fig.savefig(buf, format="png")
                    plt.close(fig)
                    graficos_simce[nivel] = base64.b64encode(buf.getvalue()).decode()
                datos["simce"] = tablas_simce
                graficos["simce"] = graficos_simce

        # RENDER HTML
        env = Environment(loader=FileSystemLoader(os.path.join(base_path, "templates")))
        template = env.get_template(template_name)
        html = template.render(
            rbd=rbd_label,
            textos=textos,
            datos=datos,
            graficos=graficos,
            resumen_territorial=resumen_data,
            fecha=hoy
        )

        st.download_button(
            label=f"⬇️ Descargar informe RBD {rbd_label}",
            data=html,
            file_name=f"informe_RBD_{rbd_label}.html",
            mime="text/html"
        )
