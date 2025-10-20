import pandas as pd
import os
import tempfile
import shutil

def leer_excel(ruta_excel: str):
    """
    Lee el Excel desde una URL o ruta local y devuelve la última fila como un diccionario.
    """
    try:
        df = pd.read_excel(ruta_excel)
        df.columns = [col.strip().upper() for col in df.columns]  # Normaliza nombres de columnas

        # ✅ Validar columnas necesarias
        columnas_necesarias = {"TICKET", "PEDIDO", "PRODUCTO", "CANT"}
        if not columnas_necesarias.issubset(set(df.columns)):
            raise ValueError(f"El archivo Excel debe contener las columnas: {', '.join(columnas_necesarias)}")

        # ✅ Tomar la última fila (la más reciente)
        ultima_fila = df.tail(1).iloc[0]

        datos = {
            "TICKET": str(ultima_fila["TICKET"]),
            "PEDIDO": str(ultima_fila["PEDIDO"]),
            "PRODUCTO": str(ultima_fila["PRODUCTO"]),
            "CANT": str(ultima_fila["CANT"]),
        }

        print(f" Último registro leído correctamente:\n  Pedido: {datos['PEDIDO']}\n  TK: {datos['TICKET']}\n  Producto: {datos['PRODUCTO']}\n  Cantidad: {datos['CANT']}")
        return datos

    except Exception as e:
        print(f"⚠️ Error al leer el archivo Excel: {e}")
        return None
