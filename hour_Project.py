import tkinter as tk
from tkinter import filedialog
import time
import pandas as pd
from datetime import datetime
import os
import shutil

def mostrar_ventana_proyectos(nombre_usuario):
    ventana_proyectos = tk.Tk()
    ventana_proyectos.title("Proyectos")

    ventana_proyectos.minsize(200, 100)

    etiqueta_bienvenida = tk.Label(ventana_proyectos, text=f"¡Bienvenido, {nombre_usuario}!")
    etiqueta_bienvenida.pack()

    file_path = None
    df = None
    proyectos = []

    def cargar_excel():
        nonlocal file_path, df, proyectos
        file_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if file_path:
            try:
                # No crear la carpeta "document", usar el directorio actual
                df = pd.read_excel(file_path)

                # Vaciar campos a partir de la columna 0 si no están vacíos
                for col in df.columns[1:]:  # Comenzar desde la columna 1
                    if not df[col].isnull().all():
                        df[col] = ""

                df["User"] = df["User"].astype(str)
                df["Start"] = df["Start"].astype(str)
                df["End"] = df["End"].astype(str)

                if "Tiempo Invertido" not in df.columns:
                    df["Tiempo Invertido"] = ""

                proyectos = df.iloc[:, 0].tolist()

                variable_proyecto.set(proyectos[0])
                desplegable_proyecto["menu"].delete(0, "end")
                for proyecto in proyectos:
                    desplegable_proyecto["menu"].add_command(label=proyecto,
                                                             command=lambda p=proyecto: cambiar_proyecto(p,
                                                                                                         nombre_usuario))

            except FileNotFoundError:
                tk.messagebox.showerror("Error", "Archivo no encontrado.")
            except pd.errors.ParserError:
                tk.messagebox.showerror("Error", "Error al leer el archivo Excel.")
            except Exception as e:
                tk.messagebox.showerror("Error", f"Error inesperado: {e}")

    boton_cargar = tk.Button(ventana_proyectos, text="Cargar Excel", command=cargar_excel)
    boton_cargar.pack()

    frame_proyecto = tk.Frame(ventana_proyectos)
    frame_proyecto.pack()

    variable_proyecto = tk.StringVar(ventana_proyectos)
    variable_proyecto.set("Seleccionar proyecto")
    desplegable_proyecto = tk.OptionMenu(frame_proyecto, variable_proyecto, "Seleccionar proyecto")
    desplegable_proyecto.pack(side=tk.LEFT)

    entrada_buscador = tk.Entry(frame_proyecto)
    entrada_buscador.pack(side=tk.LEFT)

    def buscar_proyecto(*args):
        texto_busqueda = entrada_buscador.get().lower()
        desplegable_proyecto["menu"].delete(0, "end")
        proyectos_filtrados = [proyecto for proyecto in proyectos if texto_busqueda in proyecto.lower()]  # Corrección aquí
        for proyecto in proyectos_filtrados:
            desplegable_proyecto["menu"].add_command(label=proyecto,
                                                     command=lambda p=proyecto: cambiar_proyecto(p, nombre_usuario))
        if proyectos_filtrados:
            variable_proyecto.set(proyectos_filtrados[0])  # Corrección aquí
        else:
            variable_proyecto.set("Seleccionar proyecto")  # Corrección aquí

    entrada_buscador.bind("<KeyRelease>", buscar_proyecto)

    tiempo_inicio = None
    tiempo_fin = None
    proyecto_anterior = None

    etiqueta_contador = tk.Label(ventana_proyectos, text="00:00:00", font=("Helvetica", 48))
    etiqueta_contador.pack()

    def cambiar_proyecto(proyecto_seleccionado, nombre_usuario):
        nonlocal tiempo_inicio, tiempo_fin, proyecto_anterior, df, file_path
        tiempo_fin = datetime.now()

        if tiempo_inicio and proyecto_anterior and df is not None and file_path is not None:
            try:
                indice = proyectos.index(proyecto_anterior)
                df.at[indice, "User"] = nombre_usuario
                df.at[indice, "Start"] = tiempo_inicio.strftime("%Y-%m-%d %H:%M:%S")
                df.at[indice, "End"] = tiempo_fin.strftime("%Y-%m-%d %H:%M:%S")

                tiempo_invertido = tiempo_fin - tiempo_inicio
                horas, segundos = divmod(tiempo_invertido.seconds, 3600)
                minutos, segundos = divmod(segundos, 60)
                df.at[indice, "Tiempo Invertido"] = f"{horas:02d}:{minutos:02d}"  # Corrección aquí
                df.to_excel(file_path, index=False)
            except Exception as e:
                tk.messagebox.showerror("Error", f"Error al actualizar Excel: {e}")

        tiempo_inicio = datetime.now()
        print(f"tiempo_inicio: {tiempo_inicio}")
        proyecto_anterior = proyecto_seleccionado
        variable_proyecto.set(proyecto_seleccionado)
        actualizar_contador()

    def actualizar_contador():
        nonlocal tiempo_inicio
        print("actualizar_contador() llamada")
        if tiempo_inicio:
            tiempo_transcurrido = datetime.now() - tiempo_inicio
            horas, segundos = divmod(tiempo_transcurrido.seconds, 3600)
            minutos, segundos = divmod(segundos, 60)
            print(f"tiempo_transcurrido: {tiempo_transcurrido}")
            etiqueta_contador.config(text=f"{horas:02d}:{minutos:02d}:{segundos:02d}")
            ventana_proyectos.after(1000, actualizar_contador)

    def finalizar_proceso():
        nonlocal tiempo_inicio, tiempo_fin, proyecto_anterior, df, file_path
        if tiempo_inicio and proyecto_anterior and df is not None and file_path is not None:
            try:
                indice = proyectos.index(proyecto_anterior)
                tiempo_fin = datetime.now()
                df.at[indice, "User"] = nombre_usuario
                df.at[indice, "Start"] = tiempo_inicio.strftime("%Y-%m-%d %H:%M:%S")
                df.at[indice, "End"] = tiempo_fin.strftime("%Y-%m-%d %H:%M:%S")

                tiempo_invertido = tiempo_fin - tiempo_inicio
                horas, segundos = divmod(tiempo_invertido.seconds, 3600)
                minutos, segundos = divmod(segundos, 60)
                df.at[indice, "Tiempo Invertido"] = f"{horas:02d}:{minutos:02d}"  # Corrección aquí
                df.to_excel(file_path, index=False)

                # Exportar el archivo Excel al escritorio
                fecha_actual = datetime.now().strftime("%Y-%m-%d")
                export_file_name = f"{nombre_usuario}_{fecha_actual}.xlsx"
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")  # Obtener la ruta del escritorio
                export_file_path = os.path.join(desktop_path, export_file_name)
                df.to_excel(export_file_path, index=False)

                # No borrar la carpeta "document"
                # document_path = os.path.join(os.getcwd(), "document")
                # if os.path.exists(document_path):
                #     shutil.rmtree(document_path)

            except Exception as e:
                tk.messagebox.showerror("Error", f"Error al actualizar Excel: {e}")
        ventana_proyectos.destroy()

    boton_finalizar = tk.Button(ventana_proyectos, text="Finalizar", command=finalizar_proceso)
    boton_finalizar.pack()

    def cerrar_ventana():
        nonlocal tiempo_inicio, tiempo_fin, proyecto_anterior, df, file_path
        if tiempo_inicio and proyecto_anterior and df is not None and file_path is not None:
            try:
                indice = proyectos.index(proyecto_anterior)
                tiempo_fin = datetime.now()
                df.at[indice, "User"] = nombre_usuario
                df.at[indice, "Start"] = tiempo_inicio.strftime("%Y-%m-%d %H:%M:%S")
                df.at[indice, "End"] = tiempo_fin.strftime("%Y-%m-%d %H:%M:%S")

                tiempo_invertido = tiempo_fin - tiempo_inicio
                df.at[indice, "Tiempo Invertido"] = str(tiempo_invertido)
                df.to_excel(file_path, index=False)
            except Exception as e:
                tk.messagebox.showerror("Error", f"Error al actualizar Excel: {e}")
        ventana_proyectos.destroy()
    ventana_proyectos.protocol("WM_DELETE_WINDOW", cerrar_ventana)

    ventana_proyectos.mainloop()