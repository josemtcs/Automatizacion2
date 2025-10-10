import os

def obtener_ultimo_archivo_descargado(carpeta_descargas):
    
    archivos_validos = []

    for archivo in os.listdir(carpeta_descargas):
        ruta_completa = os.path.join(carpeta_descargas, archivo)

        
        if not os.path.isfile(ruta_completa):
            continue
        if archivo.lower() in ["desktop.ini", "thumbs.db"]:
            continue

       
        if archivo.lower().endswith((".pdf", ".xlsx", ".xls", ".csv", ".docx", ".txt", ".zip", ".png", ".jpg", ".jpeg")):
            archivos_validos.append(ruta_completa)

    if not archivos_validos:
        raise FileNotFoundError("No se encontró ningún archivo válido en la carpeta de descargas.")

    
    ultimo_archivo = max(archivos_validos, key=os.path.getmtime)
    return ultimo_archivo