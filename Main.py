from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
import time
import numpy as np
import os

# Directorio actual del script
current_directory = os.path.dirname(os.path.abspath(__file__))

# Ruta de descarga de la información
download_path = os.path.join(current_directory, 'Descargas')

# Configurar las opciones de Chrome
chrome_options = Options()
chrome_options.add_experimental_option('prefs', {
    'download.default_directory': download_path,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': False
})

# Tiempo de espera entre las acciones
tiempo_espera = 5

#Ruta del archivo descargado de la SFC
ruta_actual = download_path + '/ProductosFinancieros.xls'

#URL de la SFC
url = "https://www.superfinanciera.gov.co/jsp/10085198"


# Inicializar el navegador con las opciones configuradas
driver = webdriver.Chrome(options=chrome_options)
driver.implicitly_wait(tiempo_espera)
wait = WebDriverWait(driver, tiempo_espera)

# Maximizar la ventana
driver.maximize_window()

# Abrir el enlace proporcionado
driver.get(url)

#-------------------------------------------------------------------------------------------------------------------------------
# Funciones
#-------------------------------------------------------------------------------------------------------------------------------

# 1) Tipo de entidad financiera
def seleccion_entidad_financiera(entidad = 1):
    
    #Seleccionar la lista desplegable
    wait.until(EC.element_to_be_clickable((By.ID, "financialInstitutionSearch:typeFinancialInstitution_label"))).click()
    
    #Seleccionar el tipo de entidad que es necesario
    elemento = wait.until(EC.element_to_be_clickable((By.ID, f"financialInstitutionSearch:typeFinancialInstitution_{entidad}")))
    time.sleep(1)
    driver.execute_script("arguments[0].click();", elemento)
    time.sleep(1)
    
    #Agregar toda la selección del tipo de entidad
    element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".ui-picklist-button-add-all > .ui-button-text")))
    actions = ActionChains(driver)
    actions.double_click(element).perform()

# 2) Seleccionar Fecha de descarga de la información
def seleccionar_fecha_informacion(valor = 0):
    
    #Seleccionar la lista desplegable
    wait.until(EC.element_to_be_clickable((By.ID, "financialInstitutionSearch:period_label"))).click()

    #Seleccionar el periodo que se necesite
    elemento = wait.until(EC.element_to_be_clickable((By.ID, f"financialInstitutionSearch:period_{valor}")))
    time.sleep(1)
    driver.execute_script("arguments[0].click();", elemento)
    time.sleep(1)

# 3) Seleccionar el producto
def seleccionar_producto_financiero(producto = 1):

    #Seleccionar la lista desplegable
    wait.until(EC.element_to_be_clickable((By.ID, "financialInstitutionSearch:product_label"))).click()

    #Seleccionar el producto que se necesite
    elemento = wait.until(EC.element_to_be_clickable((By.ID, f"financialInstitutionSearch:product_{producto}")))
    time.sleep(1)
    driver.execute_script("arguments[0].click();", elemento)
    time.sleep(1)

# 4) Seleccionar los servicios del producto
def seleccionar_servicio_producto(servicio = 1):

    #Seleccionar la lista desplegable
    wait.until(EC.element_to_be_clickable((By.ID, "financialInstitutionSearch:service_label"))).click()

    #Seleccionar el producto que se necesite
    elemento = wait.until(EC.element_to_be_clickable((By.ID, f"financialInstitutionSearch:service_{servicio}")))
    time.sleep(1)
    driver.execute_script("arguments[0].click();", elemento)
    time.sleep(1)

# 5) Generar reporte del servicio
def generar_reporte():

    #Buscar el boton de generar reporte
    element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#financialInstitutionSearch\\3Asearch > .ui-button-text")))
    actions = ActionChains(driver)
    actions.double_click(element).perform()
    time.sleep(1)

# 6) Descargar reporte en excel
def descargar_reporte_excel():

    #Buscar el boton de generar reporte
    element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#financialInstitutionSearch\\3A creditProductReportFormExcel\\3Aj_idt35 > .ui-button-text")))
    actions = ActionChains(driver)
    actions.double_click(element).perform()
    time.sleep(1)

# 7) Descarga de todos los productos
def funcion_descargar_productos(num, Servicios, servicios_acum, fecha = 0):

    try:

        # Cambiar al primer marco
        driver.switch_to.frame(0)

        #Proceso para generar bien la información de las entidades
        seleccionar_fecha_informacion(fecha)
        seleccion_entidad_financiera(2)
        seleccionar_producto_financiero(1)
        seleccionar_servicio_producto(1)
        generar_reporte()
        seleccion_entidad_financiera(1)

        #Posicion del producto a descargar
        posicion = np.argmax(servicios_acum > num)

        #Ciclo para recorrer todos los productos
        for i in range(posicion + 1, 26):

            #Seleccionar el producto
            seleccionar_producto_financiero(i)

            #Carlcular el servicio desde el que se tendria que empezar a descargar
            servicios_pos =  num + 1 if num + 1 <= servicios_acum[0] else num - servicios_acum[max(posicion-1, 0)] + 1

            #Ciclo para recorrer todos los servicios
            for j in range(servicios_pos, Servicios[i-1] + 1):
                
                #Seleccionar el servicio
                seleccionar_servicio_producto(j)

                #Generar el reporte y descargarlo
                generar_reporte()
                descargar_reporte_excel()

                #Crear el nombre del archivo
                nuevo_nombre = download_path + f'/Archivo_num_{i}_{j}.xls'

                # Cambiar el nombre del archivo
                os.rename(ruta_actual, nuevo_nombre)

                #Agregar 1 al contador del numero de archivos descargados
                num += 1
                #print(f"Archivos descargados: {num}")
            
            posicion += 1

    except:
        return num
    
    return num

#-------------------------------------------------------------------------------------------------------------------------------
# Variables para los productos y servicios
#-------------------------------------------------------------------------------------------------------------------------------

#Se crea un rango con los productos que se requieran de la SFC que se encuentran en el 365
Productos = range(1, 26)

#Se establecen los servicios para cada uno de los productos
Servicios = [20, 13, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 11, 11, 7, 6, 10, 2, 11, 10, 6, 6, 2, 6]

#Numero de archivos descargados en la carpeta
num = len(os.listdir(download_path))

#Acumular el numero de archivos segun la cantidad de productos descargados
servicios_acum = np.cumsum(Servicios)

#-------------------------------------------------------------------------------------------------------------------------------
# Proceso de descarga de la información
#-------------------------------------------------------------------------------------------------------------------------------

while num < max(servicios_acum):
    
    num = funcion_descargar_productos(num, Servicios, servicios_acum, 0)
    driver.switch_to.default_content()
    driver.execute_cdp_cmd("Page.stopLoading", {})
    driver.refresh()
    time.sleep(1)

driver.quit()