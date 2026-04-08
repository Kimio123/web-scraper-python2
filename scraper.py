import os
import requests
from bs4 import BeautifulSoup
import pandas as pd
# Importamos las herramientas de diseño de Excel
from openpyxl.styles import Font, PatternFill

def rastrear_precios():
    url_base = "http://books.toscrape.com/"
    print(f"Iniciando el rastreador en: {url_base}...\n")
    
    respuesta = requests.get(url_base)
    
    if respuesta.status_code != 200:
        print(f"Error al conectar. Código: {respuesta.status_code}")
        return

    soup = BeautifulSoup(respuesta.text, 'html.parser')
    productos = soup.find_all('article', class_='product_pod')
    datos_extraidos = []

    for producto in productos:
        titulo = producto.h3.a['title']
        precio = producto.find('p', class_='price_color').text
        link = url_base + producto.h3.a['href']
        
        datos_extraidos.append({
            'Nombre del Producto': titulo,
            'Precio': precio,
            'Enlace': link
        })
        print(f"Capturado: {titulo} | {precio}")

    if datos_extraidos:
        df = pd.DataFrame(datos_extraidos)
        df['Precio'] = df['Precio'].str.replace('Â£', '$')
        nombre_archivo = 'reporte_precios.xlsx'
        
        # --- INICIO MAGIA DE FORMATO ---
        # Usamos ExcelWriter para poder modificar el diseño del archivo
        with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Productos')
            
            # Seleccionamos la hoja que acabamos de crear
            worksheet = writer.sheets['Productos']
            
            # 1. Diseñar el Encabezado (Fondo verde oscuro y letra blanca en negrita)
            color_fondo = PatternFill(start_color="107C41", fill_type="solid")
            fuente_blanca = Font(bold=True, color="FFFFFF")
            
            for celda in worksheet["1:1"]: # Itera solo sobre la primera fila
                celda.fill = color_fondo
                celda.font = fuente_blanca
            
            # 2. Ajustar el ancho de las columnas automáticamente
            for columna in worksheet.columns:
                max_length = 0
                letra_columna = columna[0].column_letter # Obtiene 'A', 'B', 'C'
                
                for celda in columna:
                    try:
                        if len(str(celda.value)) > max_length:
                            max_length = len(str(celda.value))
                    except:
                        pass
                
                # Le sumamos 2 espacios extra para que respire el texto
                worksheet.column_dimensions[letra_columna].width = max_length + 2
        # --- FIN MAGIA DE FORMATO ---

        print(f"\n¡Éxito! Se extrajeron {len(datos_extraidos)} productos.")
        print(f"Los datos se han guardado con diseño en: {nombre_archivo}")
        
        # Abrimos el archivo automáticamente
        os.startfile(nombre_archivo)
        
    else:
        print("\nNo se encontró ningún producto para extraer.")

if __name__ == "__main__":
    rastrear_precios()