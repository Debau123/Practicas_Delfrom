import os 
import win32print
import win32ui
import win32con
import qrcode
from PIL import Image, ImageWin

# Función para convertir colores RGB a formato compatible con Win32
def rgb(r, g, b):
    return (r & 0xFF) | ((g & 0xFF) << 8) | ((b & 0xFF) << 16)

# Función para generar un código QR y guardarlo como imagen PNG
def generar_qr(texto, url_base, box_size):
    url = f"{url_base}{texto}"  # Crear la URL completa con el texto proporcionado
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=box_size,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img_path = f"codigo_qr_{texto}.png"
    img.save(img_path)
    return img_path

# Función para imprimir texto y QR en una etiqueta
def imprimir_texto_y_qr(texto1, texto2, url_base, logo_path, size, single_copy=False):
    p = win32print.GetDefaultPrinter()

    # Definir el desplazamiento en milímetros
    mm_a_desplazar = 10  # Incrementar el desplazamiento a la derecha
    dpi = 203  # Resolución de la impresora
    shift_x = int(mm_a_desplazar * dpi / 25.4)  # Convertir mm a píxeles

    try:
        # Abrir la impresora
        printer_handle = win32print.OpenPrinter(p)

        # Configurar la impresión
        win32print.StartDocPrinter(printer_handle, 1, ("Etiqueta", None, "RAW"))
        win32print.StartPagePrinter(printer_handle)

        # Crear contexto de impresión
        printer_dc = win32ui.CreateDC()
        printer_dc.CreatePrinterDC(p)
        printer_dc.StartDoc("Print Document")
        printer_dc.StartPage()

        # Configurar fuentes y tamaños de texto basados en el tamaño de la etiqueta
        if size == "G":
            font_large_height = 30
            font_small_height = 25
            x_info = 40 + shift_x
            x_qr = 280 + shift_x
            y_start = 10  # Reducir margen superior
            qr_size = 160
            logo_height = 70
            logo_width = 160

            # Reducir el espacio entre el texto y el logotipo
            logo_y_offset = 5  # Ajuste para reducir el espacio vertical entre el logotipo y el texto
        else:
            font_large_height = 30
            font_small_height = 30
            x1 = 40 + shift_x
            x2 = 260 + shift_x
            y_start = 10  # Reducir margen superior para las etiquetas pequeñas
            qr_size = 140
            logo_height = 40

            # Reducir el espacio entre las dos etiquetas
            column_spacing = 5  # Ajuste para reducir el espacio entre las dos etiquetas pequeñas

        font_large = win32ui.CreateFont({
            "name": "Arial",
            "height": font_large_height,
            "weight": win32con.FW_BOLD
        })
        font_small = win32ui.CreateFont({
            "name": "Arial",
            "height": font_small_height,
            "weight": win32con.FW_NORMAL
        })
        printer_dc.SelectObject(font_large)
        printer_dc.SetTextAlign(win32con.TA_LEFT | win32con.TA_TOP)
        printer_dc.SetTextColor(rgb(0, 0, 0))  # Negro
        printer_dc.SetBkColor(rgb(255, 255, 255))  # Blanco

        # Imprimir etiqueta grande
        if size == "G":
            # Cargar y dibujar el logotipo
            logo = Image.open(logo_path)
            logo = logo.resize((logo_width, logo_height))
            dib_logo = ImageWin.Dib(logo)
            logo_x = x_info
            logo_y = y_start  # Ajustar el margen superior
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
            info_y = logo_y + logo_height + logo_y_offset  # Reducir el espacio entre el logotipo y el texto
            for line in additional_info:
                printer_dc.TextOut(info_x, info_y, line)
                info_y += 25  # Espacio entre líneas reducido

            # Imprimir el QR a la derecha
            qr_image_path = generar_qr(texto1, url_base, box_size=6)
            qr_img = Image.open(qr_image_path)
            qr_img = qr_img.resize((qr_size, qr_size))
            dib = ImageWin.Dib(qr_img)
            qr_x = x_qr
            qr_y = y_start
            dib.draw(printer_dc.GetHandleOutput(), (qr_x, qr_y, qr_x + qr_img.width, qr_y + qr_img.height))

            # Imprimir el texto centrado justo debajo del QR
            text_width = printer_dc.GetTextExtent(texto1)[0]
            text_x = qr_x + (qr_img.width - text_width) // 2
            printer_dc.TextOut(text_x, qr_y + qr_img.height + 10, texto1)

        else:
            # Coordenadas para la impresión de etiqueta pequeña
            # Cargar y dibujar el logotipo sobre el primer QR
            logo = Image.open(logo_path)
            logo = logo.resize((80, logo_height))  # Ajustar tamaño del logotipo
            dib_logo = ImageWin.Dib(logo)
            logo_x1 = x1 + (qr_size - logo.width) // 2
            logo_y1 = y_start
            dib_logo.draw(printer_dc.GetHandleOutput(), (logo_x1, logo_y1, logo_x1 + logo.width, logo_y1 + logo.height))

            # Imprimir el primer QR en la primera columna
            qr_image_path1 = generar_qr(texto1, url_base, box_size=10)
            qr_img1 = Image.open(qr_image_path1)
            qr_img1 = qr_img1.resize((qr_size, qr_size))  # Ajustar tamaño del QR
            dib1 = ImageWin.Dib(qr_img1)
            qr_x1 = x1
            qr_y1 = y_start + logo_height + 10  # Posición ajustada para el QR
            dib1.draw(printer_dc.GetHandleOutput(), (qr_x1, qr_y1, qr_x1 + qr_img1.width, qr_y1 + qr_img1.height))

            # Imprimir el número en vertical a la derecha del primer QR
            text_width1 = printer_dc.GetTextExtent(texto1)[0]
            text_height1 = printer_dc.GetTextExtent(texto1)[1]
            text_x1 = qr_x1 + qr_img1.width + 10
            text_y1 = qr_y1  # Ajustar la posición vertical del texto

            printer_dc.SetTextAlign(win32con.TA_RIGHT | win32con.TA_BOTTOM)
            for i, char in enumerate(texto1):
                printer_dc.TextOut(text_x1, text_y1 + (i * text_height1), char)

            if not single_copy:
                # Cargar y dibujar el logotipo sobre el segundo QR
                logo_x2 = x2 + (qr_size - logo.width) // 2
                logo_y2 = y_start
                dib_logo.draw(printer_dc.GetHandleOutput(), (logo_x2, logo_y2, logo_x2 + logo.width, logo_y2 + logo.height))

                # Imprimir el segundo QR en la segunda columna
                qr_image_path2 = generar_qr(texto2, url_base, box_size=10)
                qr_img2 = Image.open(qr_image_path2)
                qr_img2 = qr_img2.resize((qr_size, qr_size))  # Ajustar tamaño del QR
                dib2 = ImageWin.Dib(qr_img2)
                qr_x2 = x2
                qr_y2 = y_start + logo_height + column_spacing  # Reducir el espacio entre las dos etiquetas
                dib2.draw(printer_dc.GetHandleOutput(), (qr_x2, qr_y2, qr_x2 + qr_img2.width, qr_y2 + qr_img2.height))

                # Imprimir el número en vertical a la derecha del segundo QR
                text_width2 = printer_dc.GetTextExtent(texto2)[0]
                text_height2 = printer_dc.GetTextExtent(texto2)[1]
                text_x2 = qr_x2 + qr_img2.width + 10
                text_y2 = qr_y2  # Ajustar la posición vertical del texto

                for i, char in enumerate(texto2):
                    printer_dc.TextOut(text_x2, text_y2 + (i * text_height2), char)

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

        # Eliminar las imágenes temporales solo si existen
        if size == "G":
            os.remove(qr_image_path)  # Para etiquetas grandes
        else:
            os.remove(qr_image_path1)  # Para etiquetas pequeñas
            if not single_copy and size != "G":
                os.remove(qr_image_path2)

# Ruta del logotipo siempre la misma
logo_path = "Logo.png"  # Cambia esto a la ruta real de tu logotipo

# Preguntar por el tipo de etiqueta
tipo_etiqueta = input("¿Desea etiqueta grande (G) o pequeña (P)? ").upper()
while tipo_etiqueta not in ["G", "P"]:
    tipo_etiqueta = input("Entrada inválida. Por favor, ingrese 'G' para grande o 'P' para pequeña: ").upper()

# Ejecutar la función con la entrada numérica y cantidad de copias
inicio = input("Ingrese el número de inicio: ").zfill(6)
copias = int(input("Ingrese la cantidad de copias: "))
url_base = "https://admin.misepis.com/code/"

# Imprimir etiquetas grandes
if tipo_etiqueta == "G":
    for i in range(copias):
        texto = str(int(inicio) + i).zfill(6)  # Asegurarse de que tenga al menos 6 dígitos con ceros a la izquierda
        imprimir_texto_y_qr(texto, "", url_base, logo_path, size="G", single_copy=True)  # El número cambia en cada iteración
# Imprimir etiquetas pequeñas
else:
    if copias == 1:
        imprimir_texto_y_qr(inicio, "", url_base, logo_path, size="P", single_copy=True)
    else:
        for i in range(0, copias, 2):
            texto1 = str(int(inicio) + i).zfill(6)  # Asegurarse de que tenga al menos 6 dígitos con ceros a la izquierda
            texto2 = str(int(inicio) + i + 1).zfill(6)  # Asegurarse de que tenga al menos 6 dígitos con ceros a la izquierda
            imprimir_texto_y_qr(texto1, texto2, url_base, logo_path, size="P")


