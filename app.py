import os
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.styles import PatternFill
from flask import Flask, render_template, request, redirect, url_for, send_file
from tempfile import NamedTemporaryFile
from statistics import mean

# Definir los nombres de espacio XML
namespaces = {'dte': 'http://www.sat.gob.gt/dte/fel/0.2.0'}

# Crear una instancia de la aplicación Flask
app = Flask(__name__)

# Ruta para cargar el formulario HTML
@app.route('/', methods=['GET'])
def cargar_formulario():
    return render_template('formulario.html')

# Ruta para procesar archivos XML
@app.route('/procesar', methods=['POST'])
def procesar_xml():
    # Verificar si se ha enviado un archivo
    if 'archivo_xml' not in request.files:
        return "No se ha seleccionado ningún archivo XML."

    archivos = request.files.getlist('archivo_xml')

    # Verificar si se seleccionaron archivos
    if not archivos:
        return "No se han seleccionado archivos XML."

    # Crear una lista para almacenar los datos de las facturas
    datos_facturas = []

    for archivo in archivos:
        # Verificar si el archivo es válido (XML)
        if archivo.filename.endswith('.xml'):
            xml_content = archivo.read()
            resultados = extraer_datos_factura(xml_content)
            datos_facturas.append(resultados)

    if not datos_facturas:
        return "No se han encontrado archivos XML válidos."

    # Guardar los datos en un archivo Excel temporal
    excel_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
    guardar_datos_en_excel(datos_facturas, excel_file.name)

    return redirect(url_for('descargar_excel', filename=excel_file.name))

# Función para extraer datos de un archivo XML de factura
def extraer_datos_factura(xml_content):
    tree = ET.ElementTree(ET.fromstring(xml_content))
    root = tree.getroot()

    # Obtener datos generales
    datos_generales = root.find(".//dte:DatosGenerales", namespaces=namespaces)
    numero_factura = datos_generales.get("ID", "No disponible")
    fecha_emision_raw = datos_generales.get("FechaHoraEmision")
    fecha_emision = "/".join(fecha_emision_raw.split("T")[0].split("-")[::-1])
    tipo_documento = datos_generales.get("Tipo", "No disponible")

    # Obtener datos del emisor
    emisor = root.find(".//dte:Emisor", namespaces=namespaces)
    nombre_emisor = emisor.get("NombreEmisor", "No disponible")
    codigo_establecimiento = emisor.get("CodigoEstablecimiento", "No disponible")
    nit_emisor = emisor.get("NITEmisor", "No disponible")

    # Obtener datos del receptor
    receptor = root.find(".//dte:Receptor", namespaces=namespaces)
    nombre_receptor = receptor.get("NombreReceptor", "No disponible")

    # Obtener detalles de los ítems
    detalles_items = []
    items = root.findall(".//dte:Item", namespaces=namespaces)
    for item in items:
        cantidad = item.find("dte:Cantidad", namespaces=namespaces).text
        descripcion = item.find("dte:Descripcion", namespaces=namespaces).text
        bs_element = item.get("BienOServicio")
        bs = "Bien" if bs_element == "B" else "Servicio"
        precio_unitario = item.find("dte:PrecioUnitario", namespaces=namespaces).text
        total = item.find("dte:Total", namespaces=namespaces).text
        detalles_items.append([cantidad, descripcion, bs, precio_unitario, total])

    # Obtener totales
    totales = root.find(".//dte:Totales", namespaces=namespaces)
    gran_total = totales.find("dte:GranTotal", namespaces=namespaces).text if totales is not None else "No disponible"

    # Obtener impuestos
    impuestos = root.findall(".//dte:TotalImpuesto", namespaces=namespaces)
    impuestos_dict = {
        "IVA": "0",
        "PETROLEO": "0",
        "TURISMO HOSPEDAJE": "0",
        "TIMBRE DE PRENSA": "0",
        "BOMBEROS": "0",
        "BEBIDAS ALCOHOLICAS": "0",
        "BEBIDAS NO ALCOHOLICAS": "0"
    }
    for impuesto in impuestos:
        nombre_corto = impuesto.get("NombreCorto", "No disponible")
        total_monto_impuesto = impuesto.get("TotalMontoImpuesto", "0")
        if nombre_corto in impuestos_dict:
            impuestos_dict[nombre_corto] = total_monto_impuesto

    # Obtener datos de certificación
    certificacion = root.find(".//dte:Certificacion", namespaces=namespaces)
    numero_autorizacion = certificacion.find(".//dte:NumeroAutorizacion", namespaces=namespaces)
    numero_certificacion = numero_autorizacion.get("Numero") if numero_autorizacion is not None else "No disponible"
    serie_certificacion = numero_autorizacion.get("Serie") if numero_autorizacion is not None else "No disponible"

    return {
        "FechaEmision": fecha_emision,
        "TipoDocumento": tipo_documento,
        "NITEmisor": nit_emisor,
        "NombreEmisor": nombre_emisor,
        "CodigoEstablecimiento": codigo_establecimiento,
        "NombreReceptor": nombre_receptor,
        "DetallesItems": detalles_items,
        "GranTotal": gran_total,
        **impuestos_dict,
        "SerieCertificacion": serie_certificacion,
        "NumeroCertificacion": numero_certificacion
    }

# Función para guardar los datos en un archivo Excel temporal
def guardar_datos_en_excel(datos_facturas, excel_file_name):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Agregar encabezados
    encabezados = [
        "FechaEmision", "TipoDocumento", "SerieCertificacion", "NumeroCertificacion",
        "NITEmisor", "NombreEmisor", "CodigoEstablecimiento", "NombreReceptor",
        "Cantidad", "Descripcion", "B/S", "PrecioUnitario", "Total",
        "Tasa de Alumbrado Público (Cobro Municipal)", "GranTotal",
        "IVA", "PETROLEO", "TURISMO HOSPEDAJE", "TIMBRE DE PRENSA",
        "BOMBEROS", "BEBIDAS ALCOHOLICAS", "BEBIDAS NO ALCOHOLICAS"
    ]
    sheet.append(encabezados)

    # Agregar datos de facturas
    for datos_factura in datos_facturas:
        for detalle_item in datos_factura["DetallesItems"]:
            fila = [
                datos_factura["FechaEmision"],
                datos_factura["TipoDocumento"],
                datos_factura["SerieCertificacion"],
                datos_factura["NumeroCertificacion"],
                datos_factura["NITEmisor"],
                datos_factura["NombreEmisor"],
                datos_factura["CodigoEstablecimiento"],
                datos_factura["NombreReceptor"],
                detalle_item[0],  # Cantidad
                detalle_item[1],  # Descripcion
                detalle_item[2],  # B/S
                detalle_item[3],  # PrecioUnitario
                detalle_item[4],  # Total
                datos_factura["GranTotal"],
                datos_factura["IVA"],
                datos_factura["PETROLEO"],
                datos_factura["TURISMO HOSPEDAJE"],
                datos_factura["TIMBRE DE PRENSA"],
                datos_factura["BOMBEROS"],
                datos_factura["BEBIDAS ALCOHOLICAS"],
                datos_factura["BEBIDAS NO ALCOHOLICAS"]
            ]
            sheet.append(fila)

    # Guardar el archivo Excel
    workbook.save(excel_file_name)

# Ruta para descargar el archivo de Excel
@app.route('/descargar/<filename>', methods=['GET'])
def descargar_excel(filename):
    return send_file(filename, as_attachment=True, download_name='facturas.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
