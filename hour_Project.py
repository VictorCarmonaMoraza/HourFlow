import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os

def mostrar_ventana_proyectos(nombre_usuario):
    ventana_proyectos = tk.Tk()
    ventana_proyectos.title("Proyectos")

    ventana_proyectos.minsize(200, 100)

    etiqueta_bienvenida = tk.Label(ventana_proyectos, text=f"¡Bienvenido, {nombre_usuario}!")
    etiqueta_bienvenida.pack()

    file_path = None
    df = None
    proyectos = []
    datos_proyectos = []  # Lista para almacenar los datos de los proyectos
    tiempo_inicio = None
    tiempo_fin = None
    proyecto_anterior = None
    contador_pausado = False  # Variable para rastrear si el contador está pausado
    color_contador = "black"  # Color inicial del contador
    tiempo_pausa_inicio = None  # Variable para almacenar el tiempo de inicio de la pausa
    tiempo_acumulado_pausa = 0  # Variable para almacenar el tiempo acumulado durante la pausa

    def cargar_excel():
        nonlocal file_path, df, proyectos
        file_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if file_path:
            try:
                df = pd.read_excel(file_path, sheet_name="Hoja1")
                df.rename(columns={"CÓDIGO PROYECTO": "Código Proyecto", "DESCRIPCIÓN": "Descripción"}, inplace=True)

                proyectos = df["Código Proyecto"].tolist()

                variable_proyecto.set(proyectos[0])
                desplegable_proyecto["menu"].delete(0, "end")
                for proyecto in proyectos:
                    desplegable_proyecto["menu"].add_command(label=proyecto,
                                                             command=lambda p=proyecto: cambiar_proyecto(p, nombre_usuario))

            except FileNotFoundError:
                messagebox.showerror("Error", "Archivo no encontrado.")
            except pd.errors.ParserError:
                messagebox.showerror("Error", "Error al leer el archivo Excel.")
            except Exception as e:
                messagebox.showerror("Error", f"Error inesperado: {e}")

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
        proyectos_filtrados = [proyecto for proyecto in proyectos if texto_busqueda in proyecto.lower()]
        for proyecto in proyectos_filtrados:
            desplegable_proyecto["menu"].add_command(label=proyecto,
                                                     command=lambda p=proyecto: cambiar_proyecto(p, nombre_usuario))
        if proyectos_filtrados:
            variable_proyecto.set(proyectos_filtrados[0])
        else:
            variable_proyecto.set("Seleccionar proyecto")

    entrada_buscador.bind("<KeyRelease>", buscar_proyecto)

    etiqueta_contador = tk.Label(ventana_proyectos, text="00:00:00", font=("Helvetica", 48), fg=color_contador)
    etiqueta_contador.pack()

    def cambiar_proyecto(proyecto_seleccionado, nombre_usuario):
        nonlocal tiempo_inicio, tiempo_fin, proyecto_anterior, df, file_path, datos_proyectos, contador_pausado, color_contador
        tiempo_fin = datetime.now()

        if tiempo_inicio and proyecto_anterior and file_path is not None:
            try:
                indice = proyectos.index(proyecto_anterior)

                # Calcular el tiempo invertido
                tiempo_invertido = tiempo_fin - tiempo_inicio
                horas, segundos_totales = divmod(tiempo_invertido.seconds, 3600)
                minutos, segundos = divmod(segundos_totales, 60)

                # Agregar los datos del proyecto a la lista, incluyendo la descripción
                datos_proyectos.append({
                    "Código Proyecto": proyecto_anterior,
                    "Descripción": df.at[indice, "Descripción"],  # Agregar la descripción
                    "Usuario": nombre_usuario,
                    "Start": tiempo_inicio.strftime("%Y-%m-%d %H:%M:%S"),
                    "End": tiempo_fin.strftime("%Y-%m-%d %H:%M:%S"),
                    "Tiempo Invertido": f"{horas:02d}:{minutos:02d}:{segundos:02d}"
                })

            except Exception as e:
                messagebox.showerror("Error", f"Error al actualizar Excel: {e}")

        tiempo_inicio = datetime.now()
        print(f"tiempo_inicio: {tiempo_inicio}")
        proyecto_anterior = proyecto_seleccionado
        variable_proyecto.set(proyecto_seleccionado)
        contador_pausado = False  # Reiniciar el estado de pausa
        color_contador = "black"  # Reiniciar el color del contador
        etiqueta_contador.config(fg=color_contador)  # Actualizar el color del contador
        actualizar_contador()

    def actualizar_contador():
        nonlocal tiempo_inicio, contador_pausado, color_contador, tiempo_acumulado_pausa
        print("actualizar_contador() llamada")
        if tiempo_inicio and not contador_pausado:
            tiempo_transcurrido = datetime.now() - tiempo_inicio
            tiempo_total = tiempo_transcurrido.seconds + tiempo_acumulado_pausa
            horas, segundos_totales = divmod(tiempo_total, 3600)
            minutos, segundos = divmod(segundos_totales, 60)
            print(f"tiempo_transcurrido: {tiempo_transcurrido}")
            etiqueta_contador.config(text=f"{horas:02d}:{minutos:02d}:{segundos:02d}")
            ventana_proyectos.after(1000, actualizar_contador)
        elif contador_pausado and tiempo_pausa_inicio:
            tiempo_transcurrido_pausa = datetime.now() - tiempo_pausa_inicio
            horas_pausa, segundos_totales_pausa = divmod(tiempo_transcurrido_pausa.seconds, 3600)
            minutos_pausa, segundos_pausa = divmod(segundos_totales_pausa, 60)
            etiqueta_contador.config(text=f"{horas_pausa:02d}:{minutos_pausa:02d}:{segundos_pausa:02d}")
            ventana_proyectos.after(1000, actualizar_contador)

    def pausar_contador():
        nonlocal contador_pausado, color_contador, tiempo_inicio, proyecto_anterior, datos_proyectos, df, nombre_usuario, tiempo_pausa_inicio, tiempo_acumulado_pausa
        contador_pausado = not contador_pausado  # Cambiar el estado de pausa
        if contador_pausado:
            color_contador = "green"  # Cambiar el color del contador a verde
            tiempo_pausa_inicio = datetime.now()  # Guardar el tiempo actual para calcular el tiempo transcurrido
            if tiempo_inicio:
                tiempo_fin = datetime.now()
                tiempo_transcurrido = tiempo_fin - tiempo_inicio
                horas, segundos_totales = divmod(tiempo_transcurrido.seconds, 3600)
                minutos, segundos = divmod(segundos_totales, 60)

                # Agregar los datos del proyecto a la lista, incluyendo la descripción
                datos_proyectos.append({
                    "Código Proyecto": proyecto_anterior,
                    "Descripción": df.at[proyectos.index(proyecto_anterior), "Descripción"],  # Agregar la descripción
                    "Usuario": nombre_usuario,
                    "Start": tiempo_inicio.strftime("%Y-%m-%d %H:%M:%S"),
                    "End": tiempo_fin.strftime("%Y-%m-%d %H:%M:%S"),
                    "Tiempo Invertido": f"{horas:02d}:{minutos:02d}:{segundos:02d}"
                })

            tiempo_inicio = None  # Reiniciar el contador a cero
            etiqueta_contador.config(text="00:00:00")  # Mostrar "00:00:00" en el contador
            etiqueta_contador.config(fg=color_contador)  # Actualizar el color del contador

            # Cambiar el proyecto actual al proyecto especificado (23-BONP-NOP-0097-00)
            proyecto_anterior = "23-BONP-NOP-0097-00"
            variable_proyecto.set(proyecto_anterior)
        else:
            color_contador = "black"  # Cambiar el color del contador a negro
            etiqueta_contador.config(fg=color_contador)  # Actualizar el color del contador
            tiempo_inicio = datetime.now()  # Reiniciar el tiempo de inicio al reanudar
            if tiempo_pausa_inicio:
                tiempo_acumulado_pausa += (datetime.now() - tiempo_pausa_inicio).seconds
            tiempo_pausa_inicio = None #Reiniciar el tiempo de pausa
            actualizar_contador()  # Reanudar el contador

    def finalizar_proceso():
        nonlocal tiempo_inicio, tiempo_fin, proyecto_anterior, df, file_path, datos_proyectos
        if messagebox.askyesno("Confirmar",
                               "¿Estás seguro de que quieres finalizar el proceso?"):  # Popup de confirmación
            if tiempo_inicio and proyecto_anterior and file_path is not None:
                try:
                    tiempo_fin = datetime.now()
                    tiempo_invertido = tiempo_fin - tiempo_inicio
                    horas, segundos_totales = divmod(tiempo_invertido.seconds, 3600)
                    minutos, segundos = divmod(segundos_totales, 60)
                    datos_proyectos.append({
                        "Código Proyecto": proyecto_anterior,
                        "Descripción": df.at[proyectos.index(proyecto_anterior), "Descripción"],
                        "Usuario": nombre_usuario,
                        "Start": tiempo_inicio.strftime("%Y-%m-%d %H:%M:%S"),
                        "End": tiempo_fin.strftime("%Y-%m-%d %H:%M:%S"),
                        "Tiempo Invertido": f"{horas:02d}:{minutos:02d}:{segundos:02d}"
                    })
                    fecha_actual = datetime.now().strftime("%Y-%m-%d")
                    nombre_hoja = fecha_actual
                    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        pd.DataFrame(datos_proyectos).to_excel(writer, sheet_name=nombre_hoja, index=False)
                except Exception as e:
                    messagebox.showerror("Error", f"Error al actualizar Excel: {e}")
            ventana_proyectos.destroy()

    boton_pausar = tk.Button(ventana_proyectos, text="Pausar", command=pausar_contador)
    boton_pausar.pack()

    boton_finalizar = tk.Button(ventana_proyectos, text="Finalizar", command=finalizar_proceso)
    boton_finalizar.pack()

    def cerrar_ventana():
        nonlocal tiempo_inicio, tiempo_fin, proyecto_anterior, df, file_path, datos_proyectos
        if tiempo_inicio and proyecto_anterior and file_path is not None:
            try:
                tiempo_fin = datetime.now()
                tiempo_invertido = tiempo_fin - tiempo_inicio
                horas, segundos_totales = divmod(tiempo_invertido.seconds, 3600)
                minutos, segundos = divmod(segundos_totales, 60)
                datos_proyectos.append({
                    "Código Proyecto": proyecto_anterior,
                    "Descripción": df.at[proyectos.index(proyecto_anterior), "Descripción"],
                    "Usuario": nombre_usuario,
                    "Start": tiempo_inicio.strftime("%Y-%m-%d %H:%M:%S"),
                    "End": tiempo_fin.strftime("%Y-%m-%d %H:%M:%S"),
                    "Tiempo Invertido": f"{horas:02d}:{minutos:02d}:{segundos:02d}"
                })
                fecha_actual = datetime.now().strftime("%Y-%m-%d")
                nombre_hoja = fecha_actual
                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    pd.DataFrame(datos_proyectos).to_excel(writer, sheet_name=nombre_hoja, index=False)
            except Exception as e:
                messagebox.showerror("Error", f"Error al cerrar la ventana: {e}")
        ventana_proyectos.destroy()

    ventana_proyectos.protocol("WM_DELETE_WINDOW", cerrar_ventana)

    ventana_proyectos.mainloop()