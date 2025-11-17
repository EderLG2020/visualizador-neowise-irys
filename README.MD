## Autor: 
Eder Nexón Llancari Guerra
                                                                                                                               
Path                                                                                                                       
----
D:\mundomedicodental\code

## Instalacion por archivo
- Ejecutar como administrador el archivo que esta en la carpeta en la ruta output/visoriryssetup.exe y demos en siguiente pero desmarcamos la opcion de "view launch" para que no inicie aun el programa 
- Una vez que se instalo todo correctamente debemos modificar el archivo de configuracion que debe estar en la carpeta C:\Program Files\visorirys\config.json 

luego debemos editar con las rutas correspondientes 
  "carpeta_export": "C:\Program Files\visorirys\carpetalectura", // indica la crpeta a escuchar, significa la carpeta donde se va importar todos los archivos del neowise
  "folderirys": "C:\irys", // carpeta el irys normlamente se llamaria nnt

NOTA: si omitistes el paso de desactivar el "view launch" tienes que ir al administrador de tareas y buscar "visorirys.exe" luego darle en finalizar, una vez ya editadas con las rutas correspondientes en el archivo config.json nos vamos al programador de tareas y buscar "VisorIRYS Monitor" y le damos en "ejecutar"

Programador de tarea
[https://i.ibb.co/5XpKBzj2/programadortarea.png]

Installer 2 

https://i.ibb.co/HfBNTGnL/installer1.png

Installer 2
https://i.ibb.co/Ld2QQw3P/installer2.png


Installer 3
https://i.ibb.co/sJ2zG5jL/installer3.png

Installer 4
Aqui debemos desactivar la opcione

https://i.ibb.co/d4GD6Ktf/installer4.png



---

## Instalación por archivo ejecutable

1. Ejecutar como administrador el archivo que está en `output/visoriryssetup.exe`.
2. Desmarcar la opción **"View Launch"** para que el programa no se inicie inmediatamente (Recurso visual en archivo 4).
3. Una vez instalado, editar el archivo de configuración ubicado en:

```
C:\Program Files\visorirys\config.json
```

Modificar con las rutas correspondientes:

```
{
    "carpeta_export": "C:\\Program Files\\visorirys\\carpetalectura", 
    "folderirys": "C:\\irys"
}
```

> Nota: `"carpeta_export"` indica la carpeta a escuchar, donde se importarán todos los archivos del Neowise.  
> `"folderirys"` es la carpeta donde se encuentra el software iRYS (normalmente llamada `nnt`).

4. Si olvidaste desactivar “View Launch”, abre el **Administrador de tareas**, busca `visorirys.exe` y finaliza el proceso.  
5. En el **Programador de tareas**, buscar **"VisorIRYS Monitor"** y hacer clic en **Ejecutar**.

---

## Instalación para desarrollo

- **Python:** 3.10.6  
- **pip:** 22.2.2  

### Instalar paquetes necesarios
```
pip install -r requirements.txt
```

### Editar configuración
- Editar las rutas en `config.json` según tu entorno.
- Modificar `CONFIG_DIR` en `visorirys.py` con la ruta de tu entorno de desarrollo.

### Iniciar proyecto
```
python visorirys.py
```

---

## Generar archivo ejecutable

1. Instalar PyInstaller:
```
pip install pyinstaller
```

2. Crear ejecutable sin consola:
```
pyinstaller --onefile --noconsole visorirys.py
```

3. Crear ejecutable y agregar tarea programada:
```
pyinstaller --onefile --noconsole --name visorirys visorirys.py
schtasks /Create /TN "VisorIrysMonitor" /TR "C:\visorirys\visorirys.exe" /SC ONLOGON /RL HIGHEST /F
```

---

## Configuración de carpetas

- **Carpeta de eventos:**
```
D:\mundomedicodental\code\carpetalectura\
```
- **Carpeta iRYS:**
```
D:\irys
```

---

## Recursos visuales

### Instalador

1. Instalador con permisos de administrador![Installer 1](https://i.ibb.co/HfBNTGnL/installer1.png)  
2. Crear carpeta del archivo ejecutable ![Installer 2](https://i.ibb.co/Ld2QQw3P/installer2.png)  
3. Instalar![Installer 3](https://i.ibb.co/sJ2zG5jL/installer3.png)  
4. Desactivar "View Launch". ![Installer 4 (desactivar "View Launch")](https://i.ibb.co/d4GD6Ktf/installer4.png)  


### Tarea creada (Programado de tarea)
![Programador de tareas](https://i.ibb.co/5XpKBzj2/programadortarea.png)  

### Demo
![VisorIRYS](https://i.ibb.co/k6X5yQjJ/visorirys.gif)

---

## Requisitos (requirements.txt)
```
pyperclip
pyautogui
watchdog
pygetwindow
```

> Opcional: agregar `Pillow` si quieres asegurar compatibilidad con pyautogui.
