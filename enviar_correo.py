import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from config import RUTA_PERFIL, NOMBRE_PERFIL, OUTLOOK_URL, DESTINATARIO, ASUNTO, CUERPO_HTML, CARPETA_DESCARGAS,CC
import pyautogui
from utils import obtener_ultimo_archivo_descargado
import ssl

ssl._create_default_https_context = ssl._create_unverified_context


def enviar_correo():
    print("🔵 Iniciando Outlook Web...")

    chrome_options = Options()
    chrome_options.add_argument(f"user-data-dir={RUTA_PERFIL}")       
    chrome_options.add_argument(f"profile-directory={NOMBRE_PERFIL}") 
    #chrome_options.add_experimental_option("detach", True)            # No cierra el navegador al finalizar
    chrome_options.add_argument("--start-maximized")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 25)

    driver.get(OUTLOOK_URL)
    time.sleep(5) 

    try:
        # ---------- Paso 1: Botón "Correo nuevo" ----------
        nuevo_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@aria-label='Correo nuevo' or @aria-label='New mail']")
        ))
        nuevo_btn.click()
        print("✉️ Correo nuevo iniciado...")

        # ---------- Paso 2: Campo “Para” ----------
        para_input = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@role='textbox' and (@aria-label='Para' or @aria-label='To')]")
        ))
        para_input.click()
        para_input.send_keys(DESTINATARIO)
        para_input.send_keys(Keys.TAB)

                # ---------- Paso 2: Campo “CC” ----------
        CC_input = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@role='textbox' and (@aria-label='CC' or @aria-label='CC')]")
        ))
        CC_input.click()
        CC_input.send_keys(CC)
        CC_input.send_keys(Keys.ENTER)
       

        # ---------- Paso 3: Campo “Asunto” ----------
        asunto_input = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@type='text' and (@aria-label='Asunto' or @aria-label='Subject')]")
        ))
        asunto_input.click()
        asunto_input.clear()
        asunto_input.send_keys(ASUNTO)

        # ---------- Paso 4: Campo “Cuerpo del mensaje” ----------
        cuerpo_input = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@role='textbox' and (@aria-label='Cuerpo del mensaje' or @aria-label='Message body')]")
        ))
        driver.execute_script("arguments[0].innerHTML = arguments[1];", cuerpo_input, CUERPO_HTML)
        cuerpo_input.click()
        driver.execute_script("arguments[0].innerHTML = arguments[1];", cuerpo_input, CUERPO_HTML)
        print("📝 Cuerpo del correo agregado...")



        # ---------- Paso 5: Adjuntar el último archivo descargado ----------
        archivo = obtener_ultimo_archivo_descargado(CARPETA_DESCARGAS)
        print(f"📎 Último archivo a adjuntar: {archivo}")

         # --- Clic en el botón "Adjuntar archivo" ---
        
        adjuntar_btn = wait.until(EC.element_to_be_clickable((By.XPATH,"//button[contains(@aria-label,'Adjuntar archivo') and @data-automation-type='RibbonFlyoutAnchor']"
        )))
        driver.execute_script("arguments[0].scrollIntoView(true);", adjuntar_btn)
        time.sleep(1)
        adjuntar_btn.click()
        print("📎 Se hizo clic en el botón 'Adjuntar archivo'.")

        time.sleep(2)

        # --- Clic en el botón "Examinar este equipo" ---
        examinar_btn = wait.until(EC.element_to_be_clickable((By.XPATH,"//button[@name='Examinar este equipo' and @aria-label='Examinar este equipo']")))
        driver.execute_script("arguments[0].scrollIntoView(true);", examinar_btn)
        time.sleep(1)
        examinar_btn.click()
        print("🗂️ Se hizo clic en 'Examinar este equipo'.")

        time.sleep(3)

        # --- Simular la selección del archivo con pyautogui ---
        
        pyautogui.write(archivo)
        pyautogui.press("enter")
        print("📤 Archivo seleccionado correctamente.")
        
        time.sleep(5)

        
        nombre_archivo = os.path.basename(archivo)

        try:
            wait.until(EC.presence_of_element_located((
                By.XPATH,
                f"//div[@title='{nombre_archivo}' or contains(text(), '{nombre_archivo}')]"
            )))
            print(f"✅ Archivo '{nombre_archivo}' adjuntado correctamente.")
        except Exception:
            print(f"⚠️ No se pudo confirmar visualmente que '{nombre_archivo}' fue adjuntado.")
      

    # --- Confirmación antes de enviar ---
        confirmacion = input("\n❓ ¿Deseas enviar el correo ahora? (S/N): ").strip().lower()
        if confirmacion != "s":
            print("🚫 Envío cancelado por el usuario.")
            return

            # ---------- Paso 6: Enviar ----------
        enviar_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@aria-label='Enviar' or @aria-label='Send']")
        ))
        enviar_btn.click()

        print("✅ Correo enviado correctamente.")


    finally:
        print("Proceso finalizado ✅")
        
