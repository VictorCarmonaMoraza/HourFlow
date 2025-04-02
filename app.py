import tkinter as tk
from tkinter import messagebox
import hour_Project

def iniciar_sesion():
    usuario = entrada_usuario.get()

    if usuario:
        ventana_inicio_sesion.destroy()
        hour_Project.mostrar_ventana_proyectos(usuario)  # Llama a la función del nuevo archivo
    else:
        messagebox.showerror("Error", "Por favor, ingresa tu nombre de usuario.")

ventana_inicio_sesion = tk.Tk()
ventana_inicio_sesion.title("Iniciar Sesión")

etiqueta_usuario = tk.Label(ventana_inicio_sesion, text="Nombre de usuario:")
etiqueta_usuario.pack()

entrada_usuario = tk.Entry(ventana_inicio_sesion)
entrada_usuario.pack()

boton_inicio_sesion = tk.Button(ventana_inicio_sesion, text="Iniciar Sesión", command=iniciar_sesion)
boton_inicio_sesion.pack()

ventana_inicio_sesion.mainloop()