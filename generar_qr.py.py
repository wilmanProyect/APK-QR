import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import os

def generar_qr_con_texto(data, output_path, code):
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
    nuevo_alto = alto + 80  # Espacio para el texto
    img_con_texto = Image.new('RGB', (ancho, nuevo_alto), 'white')
    img_con_texto.paste(img, (0, 0))
    
    draw = ImageDraw.Draw(img_con_texto)
    font = ImageFont.load_default()
    
    # Dibujar el primer texto (ID)
    text_bbox = draw.textbbox((0, 0), code, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]
    x = (ancho - text_width) / 2
    y = alto + 5
    draw.text((x, y), code, font=font, fill='#73BF43')
    
    # Añadir el texto adicional debajo del ID con espacio
    text2 = "Fundacion Natura Bolivia"
    text2_bbox = draw.textbbox((0, 0), text2, font=font)
    text2_width = text2_bbox[2] - text2_bbox[0]
    x2 = (ancho - text2_width) / 2
    y2 = y + text_height + 20  # Añadir un espacio claro entre las líneas
    draw.text((x2, y2), text2, font=font, fill='#73BF43')
    
    img_con_texto.save(output_path)

def procesar_excel_y_generar_qrs(excel_path, hoja, columna, fila_inicio, fila_fin, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    df = pd.read_excel(excel_path, sheet_name=hoja, header=None)  # leer sin encabezados
    print("Columnas disponibles en el DataFrame:", df.columns)
    
    # Obtener la comunidad y el beneficiario
    comunidad = df.iloc[3, 3]  # D4 -> fila 4, columna 4 (índice 3,3)
    beneficiario = df.iloc[4, 3]  # D5 -> fila 5, columna 4 (índice 4,3)
    
    for i in range(fila_inicio-1, fila_fin):
        # Obtener datos de cada fila
        id_code = str(df.iloc[i, columna])
        gps1 = str(df.iloc[i, 3])
        gps2 = str(df.iloc[i, 4])
        
        # Generar el texto que irá en el QR
        qr_text = (f"ID: {id_code}\n"
                   f"Beneficiario: {beneficiario}\n"
                   f"Comunidad: {comunidad}\n"
                   f"Ubicacion:\n"
                   f"       GPS 1: {gps1}\n"
                   f"       GPS 2: {gps2}")
        
        filename = os.path.join(output_dir, f"{id_code}.png")  # Guardar con el nombre del ID
        generar_qr_con_texto(qr_text, filename, id_code)
        print(f"QR para {id_code} guardado en {filename}")

if __name__ == "__main__":
    excel_path = r"D:\MyWork\APK QR\planilla de almendras para QR.xlsx"  # Reemplaza con la ruta a tu archivo Excel
    hoja = "Chirimoya"  # Reemplaza con el nombre de la hoja
    columna = 7  # Reemplaza con el índice de la columna de los códigos (0 para A, 1 para B, etc.)
    fila_inicio = 10  # Fila de inicio
    fila_fin = 12  # Fila de fin
    output_dir = r"D:\MyWork\APK QR\qr_generados"  # Carpeta donde se guardarán los QR

    procesar_excel_y_generar_qrs(excel_path, hoja, columna, fila_inicio, fila_fin, output_dir)
