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
        box_size=6,  # Tamaño del cuadro para un QR más pequeño
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
def imprimir_texto_y_qr(texto, url_base, logo_path):
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

        # Configurar fuentes y tamaños de texto
        font_large = win32ui.CreateFont({
            "name": "Arial",
            "height": 30,  # Tamaño de la fuente ajustado
            "weight": win32con.FW_BOLD
        })
        font_small = win32ui.CreateFont({
            "name": "Arial",
            "height": 25,  # Tamaño de la fuente ajustado
            "weight": win32con.FW_NORMAL
        })

        # Seleccionar la fuente grande
        printer_dc.SelectObject(font_large)
        printer_dc.SetTextAlign(win32con.TA_LEFT | win32con.TA_TOP)
        printer_dc.SetTextColor(rgb(0, 0, 0))  # Negro
        printer_dc.SetBkColor(rgb(255, 255, 255))  # Blanco

        # Coordenadas para la impresión
        x_info = 40  # Márgenes izquierdo para la información
        x_qr = 280  # Márgenes izquierdo para el QR ajustado más al centro
        y_start = 30  # Márgenes superiores ajustados para dejar espacio para el logotipo y QR
        qr_size = 160  # Ajustar tamaño del QR
        logo_height = 70
        logo_width = 160  # Alargar el logotipo

        # Cargar y dibujar el logotipo a la izquierda
        logo = Image.open(logo_path)
        logo = logo.resize((logo_width, logo_height))  # Ajustar tamaño del logotipo
        dib_logo = ImageWin.Dib(logo)
        logo_x = x_info
        logo_y = y_start
        dib_logo.draw(printer_dc.GetHandleOutput(), (logo_x, logo_y, logo_x + logo.width, logo_y + logo.height))

        # Imprimir la información adicional a la izquierda
        printer_dc.SelectObject(font_small)
        additional_info = [
            "misepis.com",
            "info@misepis.com",
            "964 59 14 94",
            "By Salvador S.E.A"
        ]
        info_x = x_info
        info_y = logo_y + logo_height + 10  # Posición vertical debajo del logotipo
        for line in additional_info:
            printer_dc.TextOut(info_x, info_y, line)
            info_y += 30  # Espacio entre líneas

        # Imprimir el QR a la derecha
        qr_image_path = generar_qr(texto, url_base)
        qr_img = Image.open(qr_image_path)
        qr_img = qr_img.resize((qr_size, qr_size))  # Ajustar tamaño del QR
        dib = ImageWin.Dib(qr_img)
        qr_x = x_qr
        qr_y = y_start
        dib.draw(printer_dc.GetHandleOutput(), (qr_x, qr_y, qr_x + qr_img.width, qr_y + qr_img.height))

        # Imprimir el texto centrado justo debajo del QR
        text_width = printer_dc.GetTextExtent(texto)[0]
        text_x = qr_x + (qr_img.width - text_width) // 2
        printer_dc.TextOut(text_x, qr_y + qr_img.height + 10, texto)

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
        os.remove(qr_image_path)

# Ruta del logotipo siempre la misma
logo_path = "Logo.png"  # Cambia esto a la ruta real de tu logotipo

# Ejecutar la función con la entrada numérica y cantidad de copias
inicio = input("Ingrese el número de inicio: ").zfill(6)
copias = int(input("Ingrese la cantidad de copias: "))
url_base = "https://admin.misepis.com/code/"

for i in range(copias):
    texto = str(int(inicio) + i).zfill(6)  # Asegurarse de que tenga al menos 6 dígitos con ceros a la izquierda
    imprimir_texto_y_qr(texto, url_base, logo_path)
