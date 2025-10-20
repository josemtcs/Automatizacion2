import os
import time
import pyautogui
import ssl
from seleniumbase import Driver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from config import *
from utils import obtener_ultimo_archivo_descargado


ssl._create_default_https_context = ssl._create_unverified_context


def enviar_correo(preparar=True, enviar=False, driver=None):

    try:

        if preparar:
            print("🔵 Iniciando Outlook Web...")
            driver = Driver(
                browser="chrome",
                uc=True,  
                headed=True,
                no_sandbox=True,
                incognito=False,
                user_data_dir=RUTA_PERFIL,
            )
            driver.uc_open(OUTLOOK_URL)
            driver.maximize_window()
            print("🌐 Cargando Outlook Web...")
            time.sleep(2)

        # Esperar login manual (solo la primera vez)
        

               # ---------- Paso 1: Nuevo correo ----------
            driver.click_if_visible("button[aria-label='Correo nuevo'], button[aria-label='New mail']")
            print("✉️ Correo nuevo iniciado...")
            time.sleep(1)

            # ---------- Paso 2: Campo “Para” ----------
        
            para_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[aria-label='Para'][contenteditable='true']")))
        
            para_input.click()
            time.sleep(1)
            para_input.send_keys(DESTINATARIO)
            para_input.send_keys(Keys.ENTER)
                

            # ---------- Paso 2.1: Campo “CC” ----------
            cc_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[aria-label='CC'][contenteditable='true']")))
            cc_input.click()
            cc_input.send_keys(CC)
            cc_input.send_keys(Keys.ENTER)
            cc_input.click()
            cc_input.send_keys(Keys.ENTER)
            print("📧 Destinatario y CC agregados...")

            # ---------- Paso 3: Campo “Asunto” ----------
            asunto_input =  WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='Asunto']")))
            asunto_input.click()
            asunto_input.clear()
            asunto_input.send_keys(ASUNTO)
            print("🧾 Asunto agregado...")

            # ---------- Paso 4: Campo “Cuerpo del mensaje” ----------
            cuerpo_input = driver.find_element("css selector", "div[role='textbox'][aria-label='Cuerpo del mensaje'], div[role='textbox'][aria-label='Message body']")
            driver.execute_script("arguments[0].innerHTML = arguments[1];", cuerpo_input, CUERPO_HTML)
            cuerpo_input.click()
            print("📝 Cuerpo del correo agregado...")

            # ---------- Paso 5: Adjuntar el último archivo descargado ----------
            archivo = obtener_ultimo_archivo_descargado(CARPETA_DESCARGAS)
            print(f"📎 Último archivo a adjuntar: {archivo}")

            driver.click_if_visible("button[aria-label*='Adjuntar archivo']")
            print("📎 Se hizo clic en el botón 'Adjuntar archivo'.")
            time.sleep(2)
            
            driver.click_if_visible("button[name='Examinar este equipo']")
            print("🗂️ Se hizo clic en 'Examinar este equipo'.")
            time.sleep(3)

            pyautogui.write(archivo)
            pyautogui.press("enter")
            print("📤 Archivo seleccionado correctamente.")
            time.sleep(5)

            nombre_archivo = os.path.basename(archivo)
            print(f"✅ Archivo '{nombre_archivo}' adjuntado correctamente.")

        # ---------- Confirmar envío ----------
        
            print("🕐 Correo preparado, esperando confirmación del usuario...")
            return driver  # Devolvemos el navegador activo al GUI

        # ---------- Enviar correo ----------
        elif enviar and driver is not None:
            print("🚀 Enviando correo...")
            boton_enviar = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "button[aria-label='Enviar'], button[aria-label='Send']")
                )
            )
            boton_enviar.click()
            print("✅ Correo enviado correctamente.")
            time.sleep(3)
            driver.quit()
            return True
        else:
            print("❌ Parámetros inválidos para enviar_correo.")
            return False
    except Exception as e:
        print(f"⚠️ Error durante el envío: {e}")
        try:
            driver.quit()
        except:
            pass
        return False

    finally:
        print("Proceso finalizado ✅")