# Digitador Sentencias
Programas para automatizar la extracción de datos de documentos judiciales. 

## Descrición
El script *Hace_pega.py* extrae información requerida de los archivos .docx que contienen las sentencias judiciales. Esta info es rellenada en los campos correspondientes de las tablas de los documentos .docx que almacenan la información refinada de cada sentencia. El resultado queda en una nueva carpeta

El script *Copipaste.py* es útil cuando aparecen muchas sentencias parecidas donde la mayoría de los campos se repiten. Ejecutando este script se pueden copiar los campos de las tabla del archivo que ya se digitó, al resto de archivos deseados.


## Instalación
Para utilizar este programa es necesario tener instalado Python 3.8 o superior y las siguientes librerías:
* numpy
* python-docx
* docx2txt

## Ejemplo de Uso Hace_pega.py
Al comenzar a digitar una carpeta es conveniente ejecutar este script: 
python Hace_pega.py {aqui va el nombre de la carpeta que continene las sentencias} {aqui va el nombre de la carpeta con los .docx que necesitamos digitar}

```bash
python Hace_pega.py Sentencias_Judiciales D01_50
```
El resultado sería una carpeta llamada D01_50PRE (de preprocesado) donde se rellenan algunos campos donde se logró extraer información de la sentencia. También crea una carpeta Words_a_txt donde se convirtieron las sentencias de .docx a .txt, esta carpeta puede ignorarse.

## Ejemplo de Uso Copipaste.py
Cuando se digitó un archivo y seguidos de este hay muchos parecidos, conviene ejecutar: 
python Copipaste.py {aqui va el nombre de la carpeta donde estan los archivos a modificar} 
{nombre del archivo madre ya salvado (T)} {cantidad de archivos a modificar}

```bash
python Copipaste.py D01_50 2255057T 4
```

En el ejemplo anterior se está copiando la info útil del archivo 2255057T a los 4 archivos siguientes: o sea 2255058, 2255059, 2255060 y 2255061.