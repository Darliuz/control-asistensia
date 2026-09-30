from datetime import datetime
import os
import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd

# Archivos de trabajo
EXCEL_DATA = "CONSOLIDADO EPC.xlsx"
ASISTENCIA_FILE = "registro_asistencias.csv"


class AppAsistencia:

  def __init__(self, root):
    self.root = root
    self.root.title("Sistema de Registro y Control de Asistencia Diaria")
    self.root.geometry("900x700")
    self.root.config(bg="#f8f9fa")

    # Cargar base de datos del Excel
    self.cargar_datos_base()

    # Título
    titulo = tk.Label(
        root,
        text="📋 Control Diario de Asistencia - Equipo Político Comunal",
        font=("Arial", 14, "bold"),
        bg="#f8f9fa",
        fg="#0f5132",
    )
    titulo.pack(pady=15)

    # Marco de Registro
    frame_registro = tk.LabelFrame(
        root,
        text=" Registrar Asistencia por Cédula ",
        font=("Arial", 11, "bold"),
        bg="#ffffff",
        padx=15,
        pady=15,
    )
    frame_registro.pack(fill="x", padx=20, pady=5)

    lbl_cedula = tk.Label(
        frame_registro,
        text="Número de Cédula:",
        font=("Arial", 11),
        bg="#ffffff",
    )
    lbl_cedula.grid(row=0, column=0, sticky="w", padx=5)

    self.entry_cedula = tk.Entry(
        frame_registro, font=("Arial", 14), width=22, bd=2, relief="groove"
    )
    self.entry_cedula.grid(row=0, column=1, padx=10, pady=5)
    self.entry_cedula.focus()
    # Permitir registrar presionando Enter
    self.entry_cedula.bind("<Return>", lambda event: self.registrar_asistencia())

    btn_registrar = tk.Button(
        frame_registro,
        text="Registrar Asistencia",
        font=("Arial", 11, "bold"),
        bg="#198754",
        fg="white",
        padx=12,
        pady=5,
        command=self.registrar_asistencia,
    )
    btn_registrar.grid(row=0, column=2, padx=10)

    # Mensaje de estado en tiempo real
    self.lbl_resultado = tk.Label(
        root,
        text=(
            "Estado: Listo. Ingrese la cédula y presione Enter (Fecha:"
            f" {datetime.now().strftime('%Y-%m-%d')})"
        ),
        font=("Arial", 11, "italic"),
        bg="#f8f9fa",
        fg="#333",
    )
    self.lbl_resultado.pack(pady=8)

    # Marco de Consultas y Reportes
    frame_consulta = tk.LabelFrame(
        root,
        text=" Historial y Consulta por Fechas ",
        font=("Arial", 11, "bold"),
        bg="#ffffff",
        padx=15,
        pady=15,
    )
    frame_consulta.pack(fill="both", expand=True, padx=20, pady=10)

    # Filtros
    lbl_fecha = tk.Label(
        frame_consulta, text="Fecha (YYYY-MM-DD):", font=("Arial", 10), bg="#ffffff"
    )
    lbl_fecha.grid(row=0, column=0, sticky="w", padx=5, pady=5)

    self.entry_fecha_filtro = tk.Entry(
        frame_consulta, font=("Arial", 10), width=15
    )
    self.entry_fecha_filtro.grid(row=0, column=1, padx=5, pady=5)
    self.entry_fecha_filtro.insert(0, datetime.now().strftime("%Y-%m-%d"))

    btn_filtrar = tk.Button(
        frame_consulta,
        text="Ver Día",
        font=("Arial", 10, "bold"),
        bg="#0d6efd",
        fg="white",
        command=self.ver_asistencias_dia,
    )
    btn_filtrar.grid(row=0, column=2, padx=5, pady=5)

    btn_todos = tk.Button(
        frame_consulta,
        text="Ver Todo el Histórico",
        font=("Arial", 10),
        bg="#6c757d",
        fg="white",
        command=self.ver_historico,
    )
    btn_todos.grid(row=0, column=3, padx=5, pady=5)

    btn_exportar = tk.Button(
        frame_consulta,
        text="📊 Exportar Día a Excel",
        font=("Arial", 10, "bold"),
        bg="#ffc107",
        fg="#000",
        command=self.exportar_a_excel,
    )
    btn_exportar.grid(row=0, column=4, padx=10, pady=5)

    # Tabla visual (Treeview)
    columns = (
        "Cedula",
        "Nombre",
        "Municipio",
        "Parroquia",
        "Comision",
        "Fecha",
        "Hora",
    )
    self.tree = ttk.Treeview(
        frame_consulta, columns=columns, show="headings", height=12
    )

    for col in columns:
      self.tree.heading(col, text=col)
      self.tree.column(col, width=110, anchor="center")
    self.tree.column("Nombre", width=170)
    self.tree.column("Comision", width=170)

    scrollbar = ttk.Scrollbar(
        frame_consulta, orient="vertical", command=self.tree.yview
    )
    self.tree.configure(yscrollcommand=scrollbar.set)

    self.tree.grid(row=1, column=0, columnspan=5, sticky="nsew", pady=10)
    scrollbar.grid(row=1, column=5, sticky="ns", pady=10)

    frame_consulta.grid_columnconfigure(0, weight=1)
    frame_consulta.grid_rowconfigure(1, weight=1)

    # Cargar la vista inicial del día actual al abrir
    self.ver_asistencias_dia()

  def cargar_datos_base(self):
    """Carga el archivo Excel principal con la data precargada"""
    try:
      if not os.path.exists(EXCEL_DATA):
        messagebox.showerror(
            "Error", f"No se encontró el archivo '{EXCEL_DATA}' en la carpeta."
        )
        self.df_base = pd.DataFrame()
        return

      # Lee la hoja 'EQUIPO POLITICO' tomando la fila 2 como cabecera (índice 1)
      self.df_base = pd.read_excel(EXCEL_DATA, sheet_name="EQUIPO POLITICO", header=1)
      self.df_base.columns = [str(c).strip() for c in self.df_base.columns]
      self.df_base["CEDULA"] = (
          self.df_base["CEDULA"].astype(str).str.split(".").str[0].str.strip()
      )
    except Exception as e:
      messagebox.showerror(
          "Error", f"No se pudo cargar el archivo Excel correctamente: {e}"
      )
      self.df_base = pd.DataFrame()

  def registrar_asistencia(self):
    cedula = self.entry_cedula.get().strip()
    if not cedula:
      messagebox.showwarning("Atención", "Ingrese un número de cédula válido.")
      return

    # Buscar en la base de datos precargada
    persona = self.df_base[self.df_base["CEDULA"] == cedula]

    if persona.empty:
      messagebox.showerror(
          "No Encontrado",
          f"La cédula '{cedula}' no está registrada en el archivo base.",
      )
      self.entry_cedula.delete(0, tk.END)
      return

    row = persona.iloc[0]
    nombre = row.get("NOMBRE Y APELLIDO", "N/D")
    municipio = row.get("MUNICIPIO", "N/D")
    parroquia = row.get("PARROQUIA", "N/D")
    comision = row.get("COMISION", "N/D")

    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    hora_actual = datetime.now().strftime("%H:%M:%S")

    # Validar si ya registró asistencia hoy
    ya_asistio = False
    if os.path.exists(ASISTENCIA_FILE):
      df_asist = pd.read_csv(ASISTENCIA_FILE, dtype=str)
      if not df_asist.empty:
        duplicado = df_asist[
            (df_asist["CEDULA"] == cedula) & (df_asist["FECHA"] == fecha_hoy)
        ]
        if not duplicado.empty:
          ya_asistio = True

    if ya_asistio:
      self.lbl_resultado.config(
          text=(
              f"⚠️ AVISO: {nombre} (Cédula: {cedula}) ya tenía registrada su"
              f" asistencia hoy ({fecha_hoy})."
          ),
          fg="#b02a37",
      )
    else:
      # Guardar registro en el archivo CSV local
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
        df_nuevo.to_csv(ASISTENCIA_FILE, mode="a", header=False, index=False)

      self.lbl_resultado.config(
          text=(
              f"✅ ¡Asistencia Registrada! -> {nombre} ({comision}) a las"
              f" {hora_actual}"
          ),
          fg="#198754",
      )

    self.entry_cedula.delete(0, tk.END)
    self.ver_asistencias_dia()

  def ver_asistencias_dia(self):
    fecha_filtro = self.entry_fecha_filtro.get().strip()
    for item in self.tree.get_children():
      self.tree.delete(item)

    if not os.path.exists(ASISTENCIA_FILE):
      return

    df_asist = pd.read_csv(ASISTENCIA_FILE, dtype=str)
    if df_asist.empty:
      return

    if fecha_filtro:
      df_asist = df_asist[df_asist["FECHA"] == fecha_filtro]

    for _, row in df_asist.iloc[::-1].iterrows():
      self.tree.insert(
          "",
          "tk.END",
          values=(
              row.get("CEDULA"),
              row.get("NOMBRE"),
              row.get("MUNICIPIO"),
              row.get("PARROQUIA"),
              row.get("COMISION"),
              row.get("FECHA"),
              row.get("HORA"),
          ),
      )

  def ver_historico(self):
    self.entry_fecha_filtro.delete(0, tk.END)
    for item in self.tree.get_children():
      self.tree.delete(item)

    if not os.path.exists(ASISTENCIA_FILE):
      return

    df_asist = pd.read_csv(ASISTENCIA_FILE, dtype=str)
    for _, row in df_asist.iloc[::-1].iterrows():
      self.tree.insert(
          "",
          "tk.END",
          values=(
              row.get("CEDULA"),
              row.get("NOMBRE"),
              row.get("MUNICIPIO"),
              row.get("PARROQUIA"),
              row.get("COMISION"),
              row.get("FECHA"),
              row.get("HORA"),
          ),
      )

  def exportar_a_excel(self):
    fecha_filtro = self.entry_fecha_filtro.get().strip()
    if not fecha_filtro:
      messagebox.showwarning(
          "Atención",
          "Por favor escriba la fecha específica (YYYY-MM-DD) que desea"
          " exportar.",
      )
      return

    if not os.path.exists(ASISTENCIA_FILE):
      messagebox.showinfo("Información", "No hay registros de asistencia aún.")
      return

    df_asist = pd.read_csv(ASISTENCIA_FILE, dtype=str)
    df_filtrado = df_asist[df_asist["FECHA"] == fecha_filtro]

    if df_filtrado.empty:
      messagebox.showinfo(
          "Sin datos",
          f"No se encontraron asistencias registradas para la fecha"
          f" {fecha_filtro}.",
      )
      return

    nombre_archivo_salida = f"Asistencia_{fecha_filtro.replace('-', '_')}.xlsx"
    try:
      df_filtrado.to_excel(nombre_archivo_salida, index=False)
      messagebox.showinfo(
          "Éxito",
          f"Reporte exportado correctamente como:\n{nombre_archivo_salida}",
      )
    except Exception as e:
      messagebox.showerror("Error", f"No se pudo exportar el archivo: {e}")


if __name__ == "__main__":
  root = tk.Tk()
  app = AppAsistencia(root)
  root.mainloop()