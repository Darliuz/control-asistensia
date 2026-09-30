from datetime import datetime
import os
import pandas as pd
import streamlit as st

# Configuración de la página web
st.set_page_config(
    page_title="Control de Asistencia - EPC", page_icon="📋", layout="centered"
)

# Archivos de trabajo
EXCEL_DATA = "CONSOLIDADO EPC.xlsx"
ASISTENCIA_FILE = "registro_asistencias.csv"


@st.cache_data
def cargar_datos_base():
  """Carga el archivo Excel principal con la data precargada"""
  try:
    if not os.path.exists(EXCEL_DATA):
      return None
    df = pd.read_excel(EXCEL_DATA, sheet_name="EQUIPO POLITICO", header=1)
    df.columns = [str(c).strip() for c in df.columns]
    df["CEDULA"] = (
        df["CEDULA"].astype(str).str.split(".").str[0].str.strip()
    )
    return df
  except Exception as e:
    st.error(f"Error al cargar el archivo Excel base: {e}")
    return None


df_base = cargar_datos_base()

st.title("📋 Control Diario de Asistencia")
st.subheader("Equipo Político Comunal")

if df_base is None:
  st.error(
      f"⚠️️ No se encontró el archivo '{EXCEL_DATA}' en el repositorio o"
      " directorio. Por favor, súbelo para que la app funcione."
  )
else:
  # Pestañas para separar el registro público y la consulta administrativa
  tab1, tab2 = st.tabs(
      ["📝 Registrar Asistencia", "📊 Consultar e Historial"]
  )

  with tab1:
    st.markdown("### Ingrese su número de cédula para registrar su asistencia")

    with st.form(key="form_asistencia"):
      cedula_input = st.text_input(
          "Número de Cédula:", placeholder="Ej: 12345678"
      )
      submit_button = st.form_submit_button(label="Registrar Asistencia")

    if submit_button:
      cedula = cedula_input.strip()
      if not cedula:
        st.warning("Por favor, ingrese un número de cédula válido.")
      else:
        # Buscar en la base de datos
        persona = df_base[df_base["CEDULA"] == cedula]

        if persona.empty:
          st.error(
              f"La cédula '{cedula}' no se encuentra registrada en la base de"
              " datos del equipo político."
          )
        else:
          row = persona.iloc[0]
          nombre = row.get("NOMBRE Y APELLIDO", "N/D")
          municipio = row.get("MUNICIPIO", "N/D")
          parroquia = row.get("PARROQUIA", "N/D")
          comision = row.get("COMISION", "N/D")

          fecha_hoy = datetime.now().strftime("%Y-%m-%d")
          hora_actual = datetime.now().strftime("%H:%M:%S")

          # Validar duplicados hoy
          ya_asistio = False
          if os.path.exists(ASISTENCIA_FILE):
            df_asist = pd.read_csv(ASISTENCIA_FILE, dtype=str)
            if not df_asist.empty:
              duplicado = df_asist[
                  (df_asist["CEDULA"] == cedula)
                  & (df_asist["FECHA"] == fecha_hoy)
              ]
              if not duplicado.empty:
                ya_asistio = True

          if ya_asistio:
            st.warning(
                f"⚠️ ¡Aviso! {nombre} (Cédula: {cedula}) ya había registrado su"
                f" asistencia el día de hoy ({fecha_hoy})."
            )
          else:
            nuevo_reg = {
                "CEDULA": cedula,
                "NOMBRE": nombre,
                "MUNICIPIO": municipio,
                "PARROQUIA": parroquia,
                "COMISION": comision,
                "FECHA": fecha_hoy,
                "HORA": hora_actual,
            }

            df_nuevo = pd.DataFrame([nuevo_reg])
            if not os.path.exists(ASISTENCIA_FILE):
              df_nuevo.to_csv(ASISTENCIA_FILE, index=False)
            else:
              df_nuevo.to_csv(
                  ASISTENCIA_FILE, mode="a", header=False, index=False
              )

            st.success(
                f"✅ ¡Asistencia registrada con éxito!\n\n**Nombre:**"
                f" {nombre}\n**Comisión:** {comision}\n**Hora:**"
                f" {hora_actual}"
            )

  with tab2:
    st.markdown("### Historial de Asistencias Registradas")

    if not os.path.exists(ASISTENCIA_FILE):
      st.info("Aún no hay registros de asistencia guardados.")
    else:
      df_asist = pd.read_csv(ASISTENCIA_FILE, dtype=str)

      if df_asist.empty:
        st.info("No hay registros en el archivo.")
      else:
        # Filtros de fecha
        fecha_filtro = st.text_input(
            "Filtrar por Fecha (YYYY-MM-DD):",
            value=datetime.now().strftime("%Y-%m-%d"),
        )

        if fecha_filtro:
          df_filtrado = df_asist[df_asist["FECHA"] == fecha_filtro]
        else:
          df_filtrado = df_asist

        st.dataframe(df_filtrado, use_container_width=True)

        # Botón para descargar reporte en CSV/Excel rápido
        st.download_button(
            label="📥 Descargar registros filtrados en CSV",
            data=df_filtrado.to_csv(index=False).encode("utf-8"),
            file_name=f"asistencia_{fecha_filtro}.csv",
            mime="text/csv",
        )