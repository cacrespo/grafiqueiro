#
# GRAFIQUEIRO V0.1
#
# AUTOR: Carlos A. Crespo
# FECHA: febrero de 2015

# Módulos para cargar archivos XLSX / PPTX
from pptx import *
#


# Cuerpo del programa


class Grafiqueiro:

    def __init__(self):
        ''' Clase que contiene modelos, planillas y syntaxis'''
        self.ppt=None
        self.xls=None
        self.txt=""
        self.ruta_salida=""

    def abrir_ppt(self,ruta):
        '''Crea el objeto PPTX y establece texto de ruta de archivo de salida'''
        self.ppt=Presentation(ruta)
        self.ruta_salida = ruta[0:-5] + "_salida.pptx"
        pass

    def carga_shape(self,b,c,d):
        # self.ppt = Presentacion
        # b = N° de Slide
        # c = Texto
        # d = Nombre del objeto
        for j in range(0,len(self.ppt.slides[b-1].shapes)):
            if self.ppt.slides[b-1].shapes[j].name == d:
                self.ppt.slides[b-1].shapes[j].text = c
        self.ppt.save(self.ruta_salida)

def correr_syntax(ruta):
  pass

def abrir_xls(ruta):
  pass

def carga_grafico():
  pass

def carga_tabla():
  pass




# funciones : carga presentación, carga planilla de datos, carga archivo de comandos, comandos.



if __name__ == '__main__':
try:
ejecutar()
#except (KeyboardInterrupt, SystemExit):
#pass

