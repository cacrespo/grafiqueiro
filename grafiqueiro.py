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
   # a = Presentacion
   # b = N° de Slide
   # c = Texto
   # d = Nombre del objeto
    for j in range(0,len(a.slides[b-1].shapes)):
        print j
        if a.slides[b-1].shapes[j].name == d:
            a.slides[b-1].shapes[j].text = c



# funciones : carga presentación, carga planilla de datos, carga archivo de comandos, comandos.



if __name__ == '__main__':
try:
ejecutar()
#except (KeyboardInterrupt, SystemExit):
#pass

