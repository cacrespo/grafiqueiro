# -*- coding: utf-8 -*-
#
# GRAFIQUEIRO V0.1
#
# AUTOR: Carlos A. Crespo
# FECHA: febrero de 2015
# Módulos para cargar archivos XLSX / PPTX
# Dependencias: python-pptx, openpyxl

from pptx import *
from openpyxl import *

#
# Cuerpo del programa
#

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

    def abrir_xls(self,ruta):
        ''' Crea el objeto XLSX'''
        self.xls=load_workbook(ruta).get_sheet_by_name("Datos")


    def carga_shape(self,b,c,d):
        '''Carga shape con el valor de una celda
        self.ppt = Presentacion
        b = N° de Slide
        c = Texto (celda entre comillas y en mayúsculas)
        d = Nombre del objeto'''
        for j in range(0,len(self.ppt.slides[b-1].shapes)):
            if self.ppt.slides[b-1].shapes[j].name == d:
                self.ppt.slides[b-1].shapes[j].text = self.xls[c.upper()].value
        self.ppt.save(self.ruta_salida)

    def carga_grafico(self,b,c,d):
        '''Carga grafico con valores de un rango de celdas
        self.ppt = Presentacion
        b = N° de Slide
        c = Rango de celdas
        d = Nombre del objeto'''
        for j in range(0,len(self.ppt.slides[b-1].shapes)):
            if self.ppt.slides[b-1].shapes[j].name == d:
                #print self.ppt.slides[b-1].shapes[j].table
                pass

    def carga_tabla(self,b,c,d):
        '''Carga grafico con valores de un rango de celdas
        self.ppt = Presentacion
        b = N° de Slide
        c = Rango de celdas
        d = Nombre del objeto'''
        pass

    def carga_caja():
        pass





    def correr_syntax(ruta):
        pass






# funciones : carga presentación, carga planilla de datos, carga archivo de comandos, comandos.



if __name__ == '__main__':
try:
ejecutar()
#except (KeyboardInterrupt, SystemExit):
#pass

