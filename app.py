from flask import Flask, request, render_template, session, jsonify
from selenium.webdriver import ActionChains
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from openpyxl import Workbook

app = Flask(__name__)

@app.route('/')
def prueba():
	wb = Workbook()
	ruta = './static/centros.xlsx'
	hoja = wb.active
	hoja.title = "Números centros dentales"
	centros = []
	numeros = []
	driver = webdriver.Chrome()
	driver.get("http://regcess.mscbs.es/regcessWeb/inicioBuscarCentrosAction.do")
	localidad = driver.find_element_by_css_selector("#agenteComboId")
	info1 = ActionChains(driver).click(localidad).perform()
	local = driver.find_element_by_xpath("/html/body/form/div[2]/div[3]/select/option[14]").click()
	priv = driver.find_element_by_css_selector("#codTipoDependenciaId")
	info3 = ActionChains(driver).click(priv).perform()
	privado = driver.find_element_by_css_selector("#codTipoDependenciaId > option:nth-child(4)").click()
	#info4 = ActionChains(driver).click(privado).perform()
	indice = driver.find_elements_by_css_selector("#tipoCentroId > option")
	a = 1
	terminado = 0
	selector = driver.find_element_by_css_selector("#tipoCentroId > option:nth-child(11)")
	datos = ActionChains(driver).click(selector).perform()
	datos1 = driver.find_element_by_css_selector("body > form > div.formLayout > div.formFoot > input")
	datos2 = ActionChains(driver).click(datos1).perform()
	indice2 = driver.find_elements_by_css_selector("body > div.tableContainer > table > tbody > tr")
	for x in range(1, 21):
		driver.find_element_by_link_text('siguiente>>').click()
	if terminado == 0:
		def datos(b):
			for b in range(2, len(indice2) + 1):
				datos3 = driver.find_element_by_css_selector("body > div.tableContainer > table > tbody > tr:nth-child(" + str(b) + ") > td:nth-child(9) > a")
				datos4 = ActionChains(driver).click(datos3).perform()
				b = b + 1
				datos5 = driver.find_element_by_css_selector("body > div.tableContainer > div:nth-child(7) > div.caja3 > div.campoSeccionDetalleCentro").text
				datos6 = driver.find_element_by_css_selector("body > div.tableContainer > div:nth-child(5) > div.caja1 > div.campoSeccionDetalleCentro").text
				print(datos6, datos5)
				numeros.append(datos5)
				centros.append(datos6)
				i = len(numeros)
				hoja['A' + str(i)] = datos6
				hoja['B' + str(i)] = datos5
				wb.save(filename = ruta)
				datos7 = driver.find_element_by_css_selector("body > div.tableContainer > div.tableHead > div > form > input")
				datos8 = ActionChains(driver).click(datos7).perform()

				if b == len(indice2):
					try:
						datas = driver.find_element_by_link_text('siguiente>>')

						if datas.text == "siguiente>>":
							datosxs = ActionChains(driver).click(datas).perform()
							datos(2)

					except NoSuchElementException:
						terminado = 1
						break
		datos(2)				
	driver.close()

	'''for i in range(0, len(numeros) + 1):

		hoja['A' + str(i)] = centros[i]
		hoja['B' + str(i)] = numeros[i]'''

	wb.save(filename = ruta)
	print(centros, numeros)
	return jsonify(centros, numeros)

if __name__ == '__main__':
	app.run(debug = True)