import os 
import getpass
import platform
from lector_excel import leer_excel


USUARIO = getpass.getuser()
SISTEMA = platform.system().lower()


CARPETA_USUARIO = ""
RUTA_PERFIL = ""
CARPETA_DESCARGAS = ""

if "windows"  in SISTEMA:
    CARPETA_USUARIO = f"C:\\Users\\{USUARIO}"
    RUTA_PERFIL = os.path.join(CARPETA_USUARIO, r"AppData\Local\Google\Chrome\User Data\Default")
    CARPETA_DESCARGAS = os.path.join(CARPETA_USUARIO, "Downloads")

elif "linux" in SISTEMA:
    CARPERA_USUARIO = f"/home/{USUARIO}"
    RUTA_PERFIL = os.path.join(CARPERA_USUARIO, r".config/google-chrome/Default")
    CARPETA_DESCARGAS = os.path.join(CARPERA_USUARIO, "Downloads")

else:
    raise Exception("Sistema operativo no Compatible. Solo Windows y Linux son soportados.")


NOMBRE_PERFIL = "Default"


OUTLOOK_URL = "https://outlook.office.com/mail/"  


DESTINATARIO = "comprastimed1.pr@gco.com.co"

CC_LISTA = [
    "comprastimed1.pr@gco.com.co",
    "comprastimed1.pr@gco.com.co",
    "comprastimed1.pr@gco.com.co",
    "comprastimed1.pr@gco.com.co"
]
CC = "; ".join(CC_LISTA)

EXCEL_URL = os.path.join(r"C:\Users\USER\OneDrive\Documentos\Orden de compra Sap v2.xlsx")
datos_excel = leer_excel(EXCEL_URL)

if datos_excel:

    ASUNTO = f"Orden de compra # {datos_excel['PEDIDO']} - TK # {datos_excel['TICKET']} "


    CUERPO_HTML = f"""
    <p>Buenos dias Magna ,<br><br>
    Por favor proceder con la OC <b>#{datos_excel['PEDIDO']}</b>.<br><br>
    {datos_excel['CANT']} unidades de <b>{datos_excel['PRODUCTO']}</b>.<br><br>
    Saludos,<br>
    muchas gracias, quedo atento.
    </p>
    """


if __name__ == "__main__":
    print("🧑 Usuario:", USUARIO)
    print("💻 Sistema:", SISTEMA)
    print("📁 Carpeta usuario:", CARPETA_USUARIO)
    print("🌐 Perfil Chrome:", RUTA_PERFIL)
    print("⬇️ Carpeta descargas:", CARPETA_DESCARGAS)
