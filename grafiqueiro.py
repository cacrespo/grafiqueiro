#
# GRAFIQUEIRO V0.1
#
# AUTOR: Carlos A. Crespo
# FECHA: febrero de 2015

# Módulos para cargar archivos XLSX / PPTX
from pptx import *
#


# Cuerpo del programa

def abrir_ppt(ruta):
  global ppt
  ppt = Presentation(ruta)
  pass

def correr_syntax(ruta):
  pass

def abrir_xls(ruta):
  pass

def carga_grafico():
  pass

def carga_tabla():
  pass

def carga_shape():
  pass



# funciones : carga presentación, carga planilla de datos, carga archivo de comandos, comandos.



if __name__ == '__main__':
try:
ejecutar()
#except (KeyboardInterrupt, SystemExit):
#pass

