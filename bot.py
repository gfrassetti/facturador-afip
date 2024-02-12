import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import datetime



# Ruta al archivo Excel local
archivo_excel = r'C:\Users\bruno\OneDrive\Desktop\Facturacion AFIP\Facturador.xlsx'



""" Credentials """
cuit = "27321522616"
password = "FIfi180686"
"""  """
# Cargar el libro de Excel
workbook = openpyxl.load_workbook(archivo_excel, data_only=True)
sheet = workbook["Sheet5"]
hoy = datetime.datetime.now()
hoy_formateado = hoy.strftime("%d/%m/%Y")

# Obtener el número de filas y columnas
num_cols = sheet.max_column
num_rows = sheet.max_row - 1

# Crear listas para almacenar los datos de cada columna
codigo_venta_list = []
fecha_list = []
codigo_servicio_list = []
servicio_list = []
total_list = []
converted_date = []

    # Iterar sobre las filas del Excel
for index, row in enumerate(sheet.iter_rows(min_row=2, max_row=num_rows, values_only=True), start=2):
        print(f'cantidad de filas: {num_rows}')
        print(f'-------------------- Cargando factura nro:{index} ----------------------')
        codigo_venta_list.append(row[0])
        fecha_list.append(row[1])
        codigo_servicio_list.append(row[2])
        servicio_list.append(row[3])
        total_list.append(row[4])    


for i in fecha_list:
    if i is not None:
        converted_date.append(i.strftime("%d/%m/%Y") if i is not None else None)
   


def get_value(list, index, num):
    var = ""
    for index, row in enumerate(range(num)):
        if index < len(list):
            var = list[index]
        return var    
   
            

# Configurar el navegador
options = webdriver.ChromeOptions()
options.add_argument('--disable-dev-shm-usage')
options.add_argument("enable-automation")
options.add_experimental_option("detach", True)
options.add_argument("--window-size=1920,1080")
options.add_argument("--no-sandbox")
options.add_argument("--disable-extensions")
options.add_argument("--dns-prefetch-disable")
options.add_argument("--disable-gpu")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
actions = ActionChains(driver)
# URL del formulario
url_formulario = "https://auth.afip.gob.ar/contribuyente_/login.xhtml"

# Abrir el formulario en el navegador
try:
    # Captura la ventana actual
    driver.get(url_formulario)
    input_cuit = driver.find_element(By.ID, "F1:username")
    input_cuit.send_keys(cuit)
    next_btn = driver.find_element(By.ID, "F1:btnSiguiente")
    next_btn.click()
    input_password = driver.find_element(By.ID, "F1:password")
    input_password.send_keys(password)
    input_password.click()
    next_btn = driver.find_element(By.ID, "F1:btnIngresar")
    next_btn.click()

    """ DASHBOARD """
    wait = WebDriverWait(driver, 10)
    comprobantes_en_linea = wait.until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'Comprobantes en línea')]")))
    comprobantes_en_linea.click()

    time.sleep(2)
    current_window = driver.current_window_handle
    # Cambia el control a la nueva ventana que se abrió
    for window_handle in driver.window_handles:
        if window_handle != current_window:
            driver.switch_to.window(window_handle)
            break
    
    time.sleep(3)
    driver.execute_script("document.querySelector('div#encabezado_logo_afip img').src = 'https://res.cloudinary.com/practicaldev/image/fetch/s--Rr7K5gOm--/c_limit%2Cf_auto%2Cfl_progressive%2Cq_auto%2Cw_880/https://dbalas.gallerycdn.vsassets.io/extensions/dbalas/vscode-html2pug/0.0.2/1532242577062/Microsoft.VisualStudio.Services.Icons.Default'")
    driver.execute_script("document.querySelector('div#encabezado_logo_afip img').style.width = '4rem'")
    for index, row in enumerate(range(num_rows)):
        btn_empresa = driver.find_element(By.CLASS_NAME, "btn_empresa")
        time.sleep(1)
        btn_empresa.click()
        generar_comprobantes = driver.find_element(By.ID, "btn_gen_cmp").click()
        puntos_de_venta = driver.find_element(By.XPATH, "//select[@name='puntoDeVenta']").click()
        actions.send_keys(Keys.ARROW_DOWN)
        actions.perform()
        actions.send_keys(Keys.ENTER)
        actions.perform()
        actions.send_keys(Keys.TAB)
        actions.perform()
        time.sleep(2)
        actions.send_keys(Keys.TAB)
        actions.perform()    
            
        actions.send_keys(Keys.TAB)
        actions.perform()    

        actions.send_keys(Keys.ENTER)
        actions.perform()  

        fecha = driver.find_element(By.ID, "fc")
        fecha.clear()
        fecha.send_keys(get_value(converted_date, index, num_rows))
        time.sleep(2)
        actions.send_keys(Keys.TAB)
        actions.perform()    
            
        actions.send_keys(Keys.TAB)
        actions.perform()

        actions.send_keys(Keys.ARROW_DOWN)
        actions.perform()

        actions.send_keys(Keys.ARROW_DOWN)
        actions.perform()
        actions.send_keys(Keys.TAB)
        actions.perform()
        actions.send_keys(Keys.TAB)
        actions.perform()

        actions.send_keys(get_value(converted_date, index, num_rows))
        actions.perform()
        time.sleep(2)
        actions.send_keys(Keys.TAB)
        actions.perform()
        actions.send_keys(Keys.TAB)
        actions.perform()

        actions.send_keys(get_value(converted_date, index, num_rows))
        actions.perform()  

        actividad = driver.find_element(By.ID, "actiAsociadaId").click()
        actions.send_keys(Keys.ARROW_DOWN)
        actions.perform()
        actions.send_keys(Keys.ARROW_DOWN)
        actions.perform()
        actions.send_keys(Keys.ARROW_DOWN)
        actions.perform()
        actions.send_keys(Keys.ENTER)
        actions.perform()
        ref = driver.find_element(By.ID, "refComEmisor")
        ref.send_keys(get_value(codigo_venta_list, index, num_rows))
        continuar = driver.find_element(By.XPATH, "//input[@value='Continuar >']")
        continuar.click()

except Exception as e:
    print(f"Error: {str(e)}")



