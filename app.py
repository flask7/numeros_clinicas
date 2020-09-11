from flask import Flask, request, render_template, session, jsonify
from selenium.webdriver import ActionChains
from selenium import webdriver
from openpyxl import Workbook

app = Flask(__name__)

@app.route('/')
def home():
	centros = []
	numeros = []
	driver = webdriver.Chrome()
	driver.get("http://regcess.mscbs.es/regcessWeb/inicioBuscarCentrosAction.do")
	indice = driver.find_elements_by_css_selector("#tipoCentroId > option")
	a = 0
	b = 0
	for i in indice:
		b = 0
		a = a + 1
		selector = driver.find_element_by_css_selector("#tipoCentroId > option:nth-child(" + str(a) + ")")
		datos = ActionChains(driver).click(selector).perform()
		datos1 = driver.find_element_by_css_selector("body > form > div.formLayout > div.formFoot > input")
		indice2 = driver.find_elements_by_css_selector("body > div.tableContainer > table > tbody > tr")
		datos2 = ActionChains(driver).click(datos1).perform()
		for x in indice2:
			b = b + 1
			datos3 = driver.find_element_by_css_selector("body > div.tableContainer > table > tbody > tr:nth-child(" + str(b) + ") > td:nth-child(9) > a")
			datos4 = ActionChains(driver).click(datos3).perform()
			datos5 = driver.find_element_by_css_selector("body > div.tableContainer > div:nth-child(7) > div.caja3 > div.campoSeccionDetalleCentro").text
			datos6 = driver.find_element_by_css_selector("body > div.tableContainer > div:nth-child(5) > div.caja1 > div.campoSeccionDetalleCentro").text
			numeros.append(datos5)
			centros.append(datos6)
			datos7 = driver.find_element_by_css_selector("body > div.tableContainer > div.tableHead > div > form > input")
			datos8 = ActionChains(driver).click(datos7).perform()
			datos9 = driver.find_element_by_css_selector("body > div.tableContainer > div.tableHead > div > button")
			datos10 = ActionChains(driver).click(datos9).perform()
	driver.close()
	data = {datos6: datos5}
	wb = Workbook()
	ruta = './static/centros.xlsx'
	hoja = wb.active
	hoja.title = "Números centros"
	fila = 1 
	col_fecha = 1 
	col_dato = 2
	for dato in data.items():
	    hoja.cell(column=col_fecha, row=fila, value=dato[0])
	    hoja.cell(column=col_dato, row=fila, value=dato[1])
	fila+=1
	wb.save(filename = ruta)
	print(centros, numeros)
	return jsonify(centros, numeros)

if __name__ == '__main__':
	app.run(debug = True)