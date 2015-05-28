
prueba = Grafiqueiro()
prueba.abrir_ppt("C:\Users\carlosc\Desktop\Grafiqueiro\prueba.pptx")
prueba.abrir_xls("C:\Users\carlosc\Desktop\Grafiqueiro\prueba.xlsx")
grafico = prueba.identifica(1,"gra 1")
datos = ChartData()
datos.categories =  ['West', 'East', 'North', 'South', 'Other']
datos.add_series('Series 1', (0.135, 0.324, 0.180, 0.235, 0.126))
datos.add_series('Series 2', (0.5, 0.34, 0.80, 0.235, 0.126))
datos.add_series('Series 3', (0.1, 0.34, 0.80, 0.235, 0.126))
