

# Guía de Instalación del Programa de Impresión de Etiquetas en Windows

## 1. Instalar Python

1. Descarga e instala Python desde la [página oficial](https://www.python.org/downloads/).

2. Durante la instalación, selecciona la opción **"Add Python to PATH"** para que Python pueda ser utilizado desde cualquier terminal.

3. Abre una terminal (CMD o PowerShell) y ejecuta los siguientes comandos para verificar que Python y `pip` están correctamente instalados:
   ```bash
   python --version
   pip --version
   
4. Instalar las dependencias necesarias
En tu ordenador original (donde ya funciona el programa), crea un archivo requirements.txt que contenga todas las dependencias de tu proyecto. Ejecuta el siguiente comando en una terminal:

bash
pip freeze > requirements.txt
Copia el archivo requirements.txt al nuevo ordenador (puedes utilizar una memoria USB, correo electrónico o cualquier método de transferencia de archivos).

En el nuevo ordenador, abre una terminal en la carpeta donde se encuentra el archivo requirements.txt.
Ejecuta el siguiente comando para instalar todas las dependencias:
bash
pip install -r requirements.txt

5. Instalar librerías adicionales de Windows
Si tu proyecto utiliza librerías específicas de Windows, como pywin32, sigue estos pasos adicionales:

Ejecuta el siguiente comando para instalar pywin32:
bash
python -m pip install pywin32
Después, asegúrate de ejecutar el siguiente comando para completar la instalación:
bash
python -m pywin32_postinstall

6. Configurar impresoras
Asegúrate de que las impresoras (tanto para etiquetas grandes como pequeñas) estén correctamente configuradas en Windows.
Verifica que los nombres de las impresoras en el código coincidan con los nombres de las impresoras instaladas en Windows.

7. Transferir el código del programa
Copia todo el código de tu programa de impresión de etiquetas desde tu ordenador original al nuevo ordenador. Esto incluye los archivos .py, plantillas de etiquetas y cualquier recurso adicional que tu programa necesite.
Verifica que las rutas a los archivos en el código sean correctas en el nuevo entorno.

8. Configurar el servidor Flask (si es necesario)
Si tu proyecto depende de un servidor Flask, sigue estos pasos adicionales:

Instalar Flask:
bash
pip install Flask
Configurar el servidor Flask: Asegúrate de que el archivo principal de tu servidor Flask (por ejemplo, app.py) esté configurado correctamente.

Ejecutar el servidor: Para iniciar el servidor Flask, ejecuta:
bash
python app.py
Si deseas que el servidor esté disponible en toda la red local, utiliza el siguiente comando:
bash
flask run --host=0.0.0.0

9. Probar la funcionalidad
Prueba de impresión: Ejecuta una prueba de impresión desde el nuevo ordenador para asegurarte de que las etiquetas se imprimen correctamente en las impresoras configuradas.
Prueba del servidor (si aplica): Si estás utilizando Flask, asegúrate de que el servidor está funcionando correctamente accediendo a la interfaz web o a los puntos finales necesarios.

10. Configurar la ejecución automática (opcional)
Si deseas que tu programa o servidor Flask se ejecute automáticamente al iniciar Windows:
Crea un archivo .bat con los comandos necesarios para ejecutar tu programa.
Coloca el archivo .bat en la carpeta de Inicio de Windows (shell:startup), de modo que el programa se inicie automáticamente al encender el ordenador.