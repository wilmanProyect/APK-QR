import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import os

def generar_qr_con_texto(data, output_path, text):
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
    nuevo_alto = alto + 130  # Espacio para el texto
    img_con_texto = Image.new('RGB', (ancho, nuevo_alto), 'white')
    img_con_texto.paste(img, (0, 0))
    
    draw = ImageDraw.Draw(img_con_texto)
    font = ImageFont.load_default()
    
    # Dibujar el texto debajo del QR
    lines = text.split('\n')
    y_text = alto + 5
    for line in lines:
        text_bbox = draw.textbbox((0, 0), line, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        x = (ancho - text_width) / 2
        draw.text((x, y_text), line, font=font, fill='#73BF43')
        y_text += text_bbox[3] - text_bbox[1] + 5  # Añadir un espacio entre líneas
    
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
        
        # Generar el texto que irá en el QR y debajo del QR
        qr_text = (f"ID: {id_code}\n"
                   f"Beneficiario: {beneficiario}\n"
                   f"Comunidad: {comunidad}\n"
                   f"Fecha: 14 de junio del 2024\n"
                   f"Ubicacion:\n"
                   f"       Longitud: {gps1}\n"
                   f"       Latitud: {gps2}\n"
                   f"Fundacion Natura Bolivia")
        
        text_below_qr = (f"ID: {id_code}\n"
                         f"Beneficiario: {beneficiario}\n"
                         f"Comunidad: {comunidad}\n"
                         f"Fecha: 14 de junio del 2024 \n"
                         f"Ubicacion:\n"
                         f"     Longitud: {gps1}\n"
                         f"     Latitud: {gps2}\n"
                         f"Fundacion Natura Bolivia")
        
        filename = os.path.join(output_dir, f"{id_code}.png")  # Guardar con el nombre del ID
        generar_qr_con_texto(qr_text, filename, text_below_qr)
        print(f"QR para {id_code} guardado en {filename}")
if __name__ == "__main__":
    excel_path = r"D:\MyWork\APK QR\planilla de almendras para QR.xlsx"  # Reemplaza con la ruta a tu archivo Excel
    hoja = "Carmencita 3"  # Reemplaza con el nombre de la hoja
    columna = 7  # Reemplaza con el índice de la columna de los códigos (0 para A, 1 para B, etc.)
    fila_inicio = 10  # Fila de inicio
    fila_fin = 179  # Fila de fin
    output_dir = r"D:\MyWork\APK QR\qr_generados"  # Carpeta donde se guardarán los QR

    procesar_excel_y_generar_qrs(excel_path, hoja, columna, fila_inicio, fila_fin, output_dir)
