import tkinter as tk
from tkinter import filedialog
import pandas as pd

def cargar_archivos():
    archivos = filedialog.askopenfilenames(filetypes=[('Archivos de Excel', '*.xlsx')])
    # Convierte los archivos Excel en DataFrames de pandas
    dataframes = [pd.read_excel(archivo) for archivo in archivos]
    # Verifica si los colaboradores y las horas son consistentes en los tres archivos
    resultado = verificar_consistencia(dataframes)
    if resultado[0]:
        mostrar_mensaje("Los archivos son consistentes")
    else:
        mensaje = resultado[1]
        mostrar_mensaje(mensaje)

def verificar_consistencia(dataframes):
    colaboradores = set(dataframes[0]['id_colaborador'])
    inconsistencias = []
    for i in range(1, len(dataframes)):
        colaboradores_otro = set(dataframes[i]['id_colaborador'])
        colaboradores_faltantes = colaboradores - colaboradores_otro
        colaboradores_extra = colaboradores_otro - colaboradores
        if colaboradores_faltantes:
            inconsistencias.append(f"Colaboradores faltantes en el archivo {i+1}: {colaboradores_faltantes}")
        if colaboradores_extra:
            inconsistencias.append(f"Colaboradores extra en el archivo {i+1}: {colaboradores_extra}")
    for i in range(len(dataframes)):
        for j in range(i+1, len(dataframes)):
            merge_data = pd.merge(dataframes[i], dataframes[j], on='id_colaborador', suffixes=('_archivo'+str(i+1), '_archivo'+str(j+1)), how='outer')
            diferencias_horas = merge_data[merge_data['horas_trabajadas_archivo'+str(i+1)] != merge_data['horas_trabajadas_archivo'+str(j+1)]]
            if not diferencias_horas.empty:
                mensaje_diferencias = "Diferencias de horas trabajadas:"
                for _, row in diferencias_horas.iterrows():
                    mensaje_diferencias += f"\nColaborador '{row['id_colaborador']}' tiene {row['horas_trabajadas_archivo'+str(i+1)]} horas en archivo {i+1} y {row['horas_trabajadas_archivo'+str(j+1)]} horas en archivo {j+1}"
                inconsistencias.append(mensaje_diferencias)
    if inconsistencias:
        mensaje = "Se encontraron las siguientes inconsistencias:\n\n" + "\n\n".join(inconsistencias)
        return (False, mensaje)
    return (True, None)

def mostrar_mensaje(mensaje):
    ventana = tk.Toplevel()
    ventana.title("Resultado")
    etiqueta = tk.Label(ventana, text=mensaje)
    etiqueta.pack()

ventana_principal = tk.Tk()
ventana_principal.title("Control de Horas")

boton_cargar = tk.Button(ventana_principal, text="Cargar archivos", command=cargar_archivos)
boton_cargar.pack()

ventana_principal.mainloop()
