## Autor: 
Eder Nexón Llancari Guerra

## Instalar paquetes de requirements para python

```bash
pip install -r requirements.txt
```

## VERSION 

PYTHON: 3.14.0    


## PASOS:

- detectar folder de eventos
- ejecutar script para que la ui pueda visualizar


## INSTALAR PAQUETES:
pip install -r requirements.txt

## Iniciar proyecto 

```bash
python visorirys.py
```

## Se otiene error
- codigo 01.021.0034
 se puee encontrar mayor informacion en: file:///C:/Program%20Files/CFLLog/guide/Software_Errors_Guide_16_6.html#ArchiveError

indica 
|ERROR | CODE |	MESSAGE | REASON | SOLUTION | SEE
01.021.0034 | The brand configuration file is not valid or not present. Please, re-install the software	| Is not possible to verify the software's brand	| Reinstall the software	


### Problema a resolver:

- Se debe detectar solo los folder que se crean en la carpeta "carpetalectura" de ello imprimir en consola que 
el nombre de la carpetay decir "procede abrir el ejecutable" y abre irys que ya tienes la ruta de del exe


## Se observo cambios al cargar la licencia

Archivos nuevos en la segunda lista (no estaban en la primera)

AnalysiSetup.dat – 1.201 bytes
Reconstructions.dat – 30.052 bytes
SetupZoom.dat – 40 bytes

Archivos modificados (mismo nombre pero tamaño distinto)

ConfDicom.ini
    Primera lista: 1.308 bytes
    Segunda lista: 1.382 bytes
NNT.ini
    Primera lista: 587 bytes
    Segunda lista: 758 bytes
swcfg.bck
    Primera lista: 91 bytes (11:57)
    Segunda lista: 91 bytes (14:17) – el tamaño es igual, pero probablemente modificado (fecha/hora diferente)
swcfg.dat
    Primera lista: 91 bytes (11:24)
    Segunda lista: 91 bytes (14:06) – igual que el anterior, tamaño idéntico, pero fecha diferente


## EXPORTABLE

instalar
```bash
pip install pyinstaller
```
y ejecutar lo siguiente:

```bash
pyinstaller --onefile --noconsole visorirys.py
```