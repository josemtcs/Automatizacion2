import tkinter as tk
from tkinter import messagebox
import subprocess
import os
from enviar_correo import enviar_correo
from utils import obtener_ultimo_archivo_descargado
from config import CARPETA_DESCARGAS


def ejecutar_enviar_correo():
    try:
        archivo = obtener_ultimo_archivo_descargado(CARPETA_DESCARGAS)
        if not archivo:
            messagebox.showerror("Error", "❌ No se encontró ningún archivo para adjuntar.")
            return

        lbl_estado.config(text="📨 Preparando correo... por favor espera.")
        ventana.update_idletasks()

        # Primera fase: preparar
        driver = enviar_correo(preparar=True, enviar=False)

        if not driver:
            messagebox.showerror("Error", "❌ No se pudo preparar el correo.")
            lbl_estado.config(text="❌ Error al preparar el correo.")
            return

        # Confirmar envío
        confirmar = messagebox.askyesno(
            "Confirmar envío",
            f"¿Deseas enviar el correo con el siguiente archivo?\n\n📎 {archivo}",
        )

        if confirmar:
            lbl_estado.config(text="🚀 Enviando correo... por favor espera.")
            ventana.update_idletasks()

            resultado = enviar_correo(preparar=False, enviar=True, driver=driver)

            if resultado:
                messagebox.showinfo("Éxito", "✅ Correo enviado correctamente.")
                lbl_estado.config(text=f"✅ Correo enviado con el archivo:\n{archivo}")
            else:
                messagebox.showerror("Error", "❌ Fallo al enviar el correo.")
                lbl_estado.config(text="❌ Fallo al enviar el correo.")
        else:
            driver.quit()
            messagebox.showinfo("Cancelado", "🚫 Envío cancelado por el usuario.")
            lbl_estado.config(text="🚫 Envío cancelado por el usuario.")

    except Exception as e:
        messagebox.showerror("Error", f"⚠️ Error al enviar el correo:\n{e}")
        lbl_estado.config(text=f"❌ Error durante el envío: {e}")

def mostrar_ultimo_archivo():
    try:
        archivo = obtener_ultimo_archivo_descargado(CARPETA_DESCARGAS)
        lbl_estado.config(text=f"📂 Último archivo detectado:\n{archivo}")
    except Exception as e:
        lbl_estado.config(text=f"❌ Error: {e}")

def ejecutar_script(nombre_script):
    try:
        ruta_script = os.path.join(os.getcwd(), nombre_script)
        if not os.path.exists(ruta_script):
            messagebox.showerror("Error", f"El script '{nombre_script}' no existe.")
            return

        subprocess.Popen(["python", ruta_script], shell=True)
        messagebox.showinfo("Ejecutando", f"🚀 Script '{nombre_script}' en ejecución...")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar el script:\n{e}")

def salir():
    ventana.destroy()


# INTERFAZ PRINCIPAL

ventana = tk.Tk()
ventana.title("Panel de Automatización - Outlook y Otros Scripts")
ventana.geometry("700x500")
ventana.configure(bg="#f0f2f5")
ventana.resizable(False, False)


# CABECERA

titulo = tk.Label(
    ventana,
    text="⚙️ Panel de Automatización",
    font=("Segoe UI", 18, "bold"),
    bg="#f0f2f5",
    fg="#2f3640"
)
titulo.pack(pady=20)

subtitulo = tk.Label(
    ventana,
    text="Ejecuta tus scripts de automatización desde una sola ventana.",
    font=("Segoe UI", 10),
    bg="#f0f2f5",
    fg="#636e72"
)
subtitulo.pack()


# BOTONES DE ACCIÓN

frame_botones = tk.Frame(ventana, bg="#f0f2f5")
frame_botones.pack(pady=30)

btn_ver_archivo = tk.Button(
    frame_botones,
    text="📁 Ver último archivo descargado",
    command=mostrar_ultimo_archivo,
    bg="#0984e3", fg="white", width=35, height=2,
    font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2"
)
btn_ver_archivo.grid(row=0, column=0, padx=10, pady=10)

btn_enviar_correo = tk.Button(
    frame_botones,
    text="📨 Enviar correo automático",
    command=ejecutar_enviar_correo,
    bg="#00b894", fg="white", width=35, height=2,
    font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2"
)
btn_enviar_correo.grid(row=1, column=0, padx=10, pady=10)

btn_otro_script = tk.Button(
    frame_botones,
    text="🧩 Ejecutar otro script (ejemplo)",
    command=lambda: ejecutar_script("otro_script.py"),
    bg="#6c5ce7", fg="white", width=35, height=2,
    font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2"
)
btn_otro_script.grid(row=2, column=0, padx=10, pady=10)

btn_salir = tk.Button(
    ventana,
    text="❌ Salir del programa",
    command=salir,
    bg="#d63031", fg="white", width=25, height=2,
    font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2"
)
btn_salir.pack(pady=20)

# LABEL DE ESTADO

lbl_estado = tk.Label(
    ventana,
    text="Estado: Esperando acción...",
    font=("Segoe UI", 10),
    bg="#f0f2f5",
    fg="#2d3436",
    wraplength=650,
    justify="center"
)
lbl_estado.pack(pady=5)

# ----------------------------------------
ventana.mainloop()