import os
import win32print
import win32ui
import win32con
import qrcode
from PIL import Image, ImageWin

# Función para convertir colores RGB a un formato compatible con Win32
def rgb(r, g, b):
    return (r & 0xFF) | ((g & 0xFF) << 8) | ((b & 0xFF) << 16)

# Función para generar un código QR y guardarlo como imagen PNG
def generar_qr(texto, url_base):
    # Crear la URL completa con el texto proporcionado
    url = f"{url_base}{texto}"

    # Configurar el código QR
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=9,  # Tamaño del cuadro para un QR más grande
        border=4,
    )

    # Agregar datos al QR y crear la imagen
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    # Guardar la imagen como PNG
    img_path = f"codigo_qr_{texto}.png"
    img.save(img_path)
    return img_path

# Función para imprimir el texto y el código QR en la etiqueta
def imprimir_texto_y_qr(texto1, texto2, url_base, logo_path, single_copy=False):
    # Obtener la impresora predeterminada
    p = win32print.GetDefaultPrinter()

    try:
        # Abrir la impresora
        printer_handle = win32print.OpenPrinter(p)

        # Iniciar el trabajo de impresión
        win32print.StartDocPrinter(printer_handle, 1, ("Etiqueta", None, "RAW"))
        win32print.StartPagePrinter(printer_handle)

        # Crear el contexto de impresión
        printer_dc = win32ui.CreateDC()
        printer_dc.CreatePrinterDC(p)
        printer_dc.StartDoc("Print Document")
        printer_dc.StartPage()

        # Configurar fuente y tamaño de texto
        font = win32ui.CreateFont({
            "name": "Arial",
            "height": 30,  # Tamaño de la fuente ajustado
            "weight": win32con.FW_NORMAL
        })
        printer_dc.SelectObject(font)
        printer_dc.SetTextAlign(win32con.TA_LEFT | win32con.TA_TOP)
        printer_dc.SetTextColor(rgb(0, 0, 0))  # Negro
        printer_dc.SetBkColor(rgb(255, 255, 255))  # Blanco

        # Coordenadas para las columnas (izquierda y derecha)
        x1 = 40  # Margen izquierdo
        x2 = 260  # Margen izquierdo para la segunda columna
        y_start = 30  # Margen superior ajustado para dejar espacio para el logotipo y QR
        logo_height = 40
        qr_size = 110  # Tamaño del QR ajustado

        # Cargar y dibujar el logotipo sobre el primer QR
        logo = Image.open(logo_path)
        logo = logo.resize((80, logo_height))  # Ajustar tamaño del logotipo
        dib_logo = ImageWin.Dib(logo)
        logo_x1 = x1 + (qr_size - logo.width) // 2
        logo_y1 = y_start
        dib_logo.draw(printer_dc.GetHandleOutput(), (logo_x1, logo_y1, logo_x1 + logo.width, logo_y1 + logo.height))

        # Imprimir el primer QR y texto en la primera columna
        qr_image_path1 = generar_qr(texto1, url_base)
        qr_img1 = Image.open(qr_image_path1)
        qr_img1 = qr_img1.resize((qr_size, qr_size))  # Ajustar tamaño del QR
        dib1 = ImageWin.Dib(qr_img1)
        qr_x1 = x1
        qr_y1 = y_start + logo_height + 10  # Posición ajustada para el QR
        dib1.draw(printer_dc.GetHandleOutput(), (qr_x1, qr_y1, qr_x1 + qr_img1.width, qr_y1 + qr_img1.height))

        # Calcular la posición del texto para que esté centrado debajo del QR
        text_width1 = printer_dc.GetTextExtent(texto1)[0]
        text_x1 = qr_x1 + (qr_img1.width - text_width1) // 2
        printer_dc.TextOut(text_x1, qr_y1 + qr_img1.height, texto1)  # Imprimir el texto centrado justo debajo del QR

        if not single_copy:
            # Cargar y dibujar el logotipo sobre el segundo QR
            logo_x2 = x2 + (qr_size - logo.width) // 2
            logo_y2 = y_start
            dib_logo.draw(printer_dc.GetHandleOutput(), (logo_x2, logo_y2, logo_x2 + logo.width, logo_y2 + logo.height))

            # Imprimir el segundo QR y texto en la segunda columna
            qr_image_path2 = generar_qr(texto2, url_base)
            qr_img2 = Image.open(qr_image_path2)
            qr_img2 = qr_img2.resize((qr_size, qr_size))  # Ajustar tamaño del QR
            dib2 = ImageWin.Dib(qr_img2)
            qr_x2 = x2
            qr_y2 = y_start + logo_height + 10  # Posición ajustada para el QR
            dib2.draw(printer_dc.GetHandleOutput(), (qr_x2, qr_y2, qr_x2 + qr_img2.width, qr_y2 + qr_img2.height))

            # Calcular la posición del texto para que esté centrado debajo del QR
            text_width2 = printer_dc.GetTextExtent(texto2)[0]
            text_x2 = qr_x2 + (qr_img2.width - text_width2) // 2
            printer_dc.TextOut(text_x2, qr_y2 + qr_img2.height, texto2)  # Imprimir el texto centrado justo debajo del QR

        # Finalizar la página y el documento de impresión
        printer_dc.EndPage()
        printer_dc.EndDoc()

    except Exception as e:
        print(f"Error al imprimir: {e}")

    finally:
        # Cerrar el contexto de impresión
        printer_dc.DeleteDC()
        win32print.EndPagePrinter(printer_handle)
        win32print.EndDocPrinter(printer_handle)
        win32print.ClosePrinter(printer_handle)

        # Eliminar las imágenes temporales
        os.remove(qr_image_path1)
        if not single_copy:
            os.remove(qr_image_path2)

# Ruta del logotipo siempre la misma
logo_path = "Logo.png"  # Cambia esto a la ruta real de tu logotipo

# Ejecutar la función con la entrada numérica y cantidad de copias
inicio = int(input("Ingrese el número de inicio: "))
copias = int(input("Ingrese la cantidad de copias: "))
url_base = "https://admin.misepis.com/code/"

if copias == 1:
    texto = str(inicio).zfill(6)  # Asegurarse de que tenga al menos 6 dígitos con ceros a la izquierda
    imprimir_texto_y_qr(texto, "", url_base, logo_path, single_copy=True)
else:
    for i in range(0, copias, 2):
        texto1 = str(inicio + i).zfill(6)  # Asegurarse de que tenga al menos 6 dígitos con ceros a la izquierda
        texto2 = str(inicio + i + 1).zfill(6)  # Asegurarse de que tenga al menos 6 dígitos con ceros a la izquierda
        imprimir_texto_y_qr(texto1, texto2, url_base, logo_path)
