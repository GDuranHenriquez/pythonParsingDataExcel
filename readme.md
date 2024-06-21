# Parsing data.

Aplicación simple para el parseo y limpieza de data desde archivos excel preconfigurados y creación de un nuevo excel a partir de estos datos.

A continuación lo pasos para generar el projecto.

* Instalar virtualenv, pip install virtualenv
* Crear un entorno virtual en la raiz del projecto, para esto en la terminal de comando ejecutamos,  python<3.12.3> -m venv venv
* Activamos el entorno virtual, desde la terminal en la raiz de nuestro proyecto ejcutamos, \env\Scripts\activate, ajustar la ruta según el sitema operativo, (windows comand pront, env\Scripts\activate ), (windows power shell, .\env\Scripts\Activate.ps1), (unix/linux/macOS, source env/bin/activate).
* Ahora intalamo todos los packages segun el archivo requirements.txt ejecutando, pip install -r requirements.txt
* cada ves que se instale una dependencia se debe actualizar el archivo requirements.txt, para esto aplicamo, pip freeze > requirements.txt.
* Una forma de instalar una dependencia y a la ves actuazar el archivo requirements.txt, en una sola instrucción e uar el comando, ./update_requirements.sh <nueva_dependencia>, esto ejecutara el script update_requirements.sh que eta en la raiz del royecto, instalara la nueva dependencia y actualizara el archivo en cuestion.