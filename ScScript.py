"""
Script que captura Imagen, Titulo, Desscripcion y Precio de un producto de Mercado Libre
y lo guarda en un archivo excel
"""

from time import sleep
from requests_html import HTMLSession
import pandas as pd

from PIL import Image
import requests
from io import BytesIO

#URL de la busqueda de MercadoLibre
url = "https://listado.mercadolibre.com.ve/televisor#D[A:televisor]"

session = HTMLSession()

def verificarSrc(producto):
    #Si existe data-src, entonces la imagen se encuentra en data-src, de lo contrario"""
    #se encuentra en src
    try:
        imagen = producto.find(".ui-search-result-image__element ")[0].attrs["data-src"]
    except KeyError:
        imagen = producto.find(".ui-search-result-image__element ")[0].attrs["src"]
    return imagen

def toExcel(imagenes, titulos, precios):
    #Creamos un diccionario con los datos
    data = {
        "Imagene": imagenes,
        "Titulo": titulos,
        "Precio": precios
    }

    #Creamos un dataframe con el diccionario
    df = pd.DataFrame(data)

    #Guardamos el dataframe en un archivo excel
    writer = pd.ExcelWriter('ProductosML.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Productos', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Productos']

    #Creamos un formato para las imagenes
    image_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'fg_color': '#D7E4BC',
        'font_color': '#9C0006'
    })

    #Agregamos el formato a la columna de imagenes
    worksheet.set_column('A:A', 20, image_format)

    #worksheet.insert_image('A2', 'prueba.jpg', {'x_scale': 0.1, 'y_scale': 0.1})

    #Guardamos el archivo
    writer.save()

def capturar(ml):
    imagenes = []
    titulos = []
    precios = []

    # Obtenemos todos los productos
    productos = ml.html.find(".ui-search-layout__item")
    for producto in productos: #Recorremos cada producto

        imagen = verificarSrc(producto)
        #response = requests.get(imagen)
        #img = Image.open(BytesIO(response.content))
        imagenes.append(imagen)

        titulo = producto.find(".ui-search-item__title")[0].text
        titulos.append(titulo)
        precio = producto.find(".price-tag-fraction")[0].text + "$"
        precios.append(precio)
        print(imagen, titulo, precio)

    
    toExcel(imagenes, titulos, precios)


def main():
    ml = session.get(url)
    ml.html.render(sleep=1)
    capturar(ml)
   

if __name__ == "__main__":
    main()
