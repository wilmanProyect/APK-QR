import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import os
<<<<<<< HEAD

def generar_qr_con_texto(data, output_path, texto_id, texto_info):
=======
#hello to work
def generar_qr_con_texto(data, output_path, code):
>>>>>>> 79dbe2dbe6784c3b59f6e953570d66a62ffacf69
    # Crear una instancia del objeto QRCode
    qr = qrcode.QRCode(
        version=2,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=3,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color='#73BF43', back_color='white')
    
    # Añadir texto debajo del QR
    ancho, alto = img.size
    nuevo_alto = alto + 150  # Espacio para el texto
    img_con_texto = Image.new('RGB', (ancho, nuevo_alto), 'white')
    img_con_texto.paste(img, (0, 0))
    
    draw = ImageDraw.Draw(img_con_texto)
    font = ImageFont.load_default()
    
    # Dibujar el primer texto (ID)
    y = alto + 5
    draw.text((10, y), texto_info, font=font, fill='#73BF43')
    
    img_con_texto.save(output_path)

def procesar_excel_y_generar_qrs(excel_path, hoja, columna_id, fila_inicio, fila_fin, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    df = pd.read_excel(excel_path, sheet_name=hoja, header=None)  # leer sin encabezados
    print("Columnas disponibles en el DataFrame:", df.columns)
    
    comunidad = df.iloc[3, 1]  # Comunidad (celda D4 -> index 3, column 1)
    propietario = df.iloc[4, 1]  # Nombre del propietario (celda D5 -> index 4, column 1)
    
    for i in range(fila_inicio-1, min(fila_fin, len(df))):
        if i >= len(df):
            continue
        gps1 = str(df.iloc[i, 3])  # Puntos GPS 1
        gps2 = str(df.iloc[i, 4])  # Puntos GPS 2
        id = str(df.iloc[i, columna_id])  # ID
        
        if pd.isna(gps1) or pd.isna(gps2) or pd.isna(id):
            print(f"Omitiendo fila {i+1} por datos faltantes")
            continue
        
        texto_info = (f"ID: {id}\n"
                      f"Propietario: {propietario}\n"
                      f"Comunidad: {comunidad}\n"
                      f"Ubicación:\n"
                      f"    GPS 1: {gps1}\n"
                      f"    GPS 2: {gps2}")
        
        filename = os.path.join(output_dir, f"{id}.png")  # Guardar con el nombre del ID
        generar_qr_con_texto(texto_info, filename, id, texto_info)
        print(f"QR para {id} guardado en {filename}")

if __name__ == "__main__":
    excel_path = r"D:\MyWork\APK QR\planilla de almendras para QR.xlsx"  # Reemplaza con la ruta a tu archivo Excel
    hoja = "Chirimoya"  # Reemplaza con el nombre de la hoja
    columna_id = 8  # Reemplaza con el índice de la columna de los códigos (por ejemplo, 8 para la columna H)
    fila_inicio = 10  # Fila de inicio
    fila_fin = 56  # Fila de fin
    output_dir = r"D:\MyWork\APK QR\qr_generados"  # Carpeta donde se guardarán los QR

    procesar_excel_y_generar_qrs(excel_path, hoja, columna_id, fila_inicio, fila_fin, output_dir)
