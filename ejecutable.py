import pdfplumber
import os
import pandas as pd
import re
import camelot
import datetime
import tkinter
from tkinter import filedialog

class extractor:
    path = ''
    pattern_secuencial = ''
    pattern_ruc = ''
    pattern_fecha = ''
    pattern_autorizacion = ''
    pattern_sub_12 = ''
    pattern_sub_0 = ''
    pattern_total = ''
    pattern_pago = ''
    df_general_data = pd.DataFrame()
    dato = {}

    def __init__(self, path):
        self.path = path
        self.pattern_secuencial = re.compile(r'[0-9]{3}-[0-9]{3}-[0-9]{9}')
        self.pattern_ruc = re.compile(r'(R.U.C.:+ [0-9]{13})')
        self.pattern_fecha = re.compile(r'(Fecha| [0-9]{2}/[0-9]{2}/[0-9]{4})')
        self.pattern_autorizacion = re.compile(r"[0-9]{49}")
        self.pattern_sub_12 = re.compile(r'SUBTOTAL 12%+ [0-9]{1,6}\.[0-9]{2}')
        self.pattern_sub_0 = re.compile(r"SUBTOTAL 0%+ [0-9]{1,6}\.[0-9]{2}")
        self.pattern_total = re.compile(r"VALOR TOTAL+ [0-9]{1,6}\.[0-9]{2}")
        self.pattern_pago = re.compile(r'[01256789 ]{2} +[\- ]+[ACDEIMNORST]{3}')
        self.writer = pd.ExcelWriter(self.path + 'detalle_facturas' + '.xlsx', engine='xlsxwriter')
        self.df_general_data = pd.DataFrame(self.dato, columns=['FECHA', 'RUC', 'PROVEEDOR', 'SECUENCIAL', 'AUTORIZACION',
                                                      'SUBTOTAL_12', 'SUBTOTAL_0', 'TOTAL', 'PAGO'])

    def bucle(self):
        dirs = os.listdir(self.path)
        for i, file in enumerate(dirs):
            archivo, ext = os.path.splitext(file)
            if ext != '.pdf':
                print('File: ' + str(file) + str('Is not PDF'))
            else:
                with pdfplumber.open(self.path+file) as pdf:
                    pages = pdf.pages
                    page_0 = pages[0]
                    txt = page_0.extract_text()
                    secuencial = re.findall(self.pattern_secuencial, txt)
                    ruc = re.findall(self.pattern_ruc, txt)
                    fecha = re.findall(self.pattern_fecha, txt)
                    autorizacion = re.findall(self.pattern_autorizacion, txt)
                    tables = camelot.read_pdf(self.path + file, flavor="stream", pages="1")
                    proveedor = tables[0].df.iloc[8][0]
                    ruc = ruc[0].split()
                    fecha = datetime.datetime.strptime((fecha[1]).lstrip(), '%d/%m/%Y')
                    new_row = {'FECHA': fecha, 'RUC': (ruc[1]).lstrip(), 'PROVEEDOR': proveedor,'SECUENCIAL': secuencial[0],
                               'AUTORIZACION': autorizacion[0]}
                    num_pages = len(pages)
                    data = []
                    for x in range(num_pages):
                        pgs = pages[x]
                        text = pgs.extract_text()
                        if text == None:
                            print("")
                        else:
                            sub_12 = re.findall(self.pattern_sub_12, text)
                            sub_0 = re.findall(self.pattern_sub_0, text)
                            total = re.findall(self.pattern_total, text)
                            pago = re.findall(self.pattern_pago, text)
                            if len(sub_12) == 0:
                                print("")
                            else:
                                sub_12 = sub_12[0].split()
                                sub_0 = sub_0[0].split()
                                total = total[0].split()
                                new_row['SUBTOTAL_12'] = sub_12[2]
                                new_row['SUBTOTAL_0'] = sub_0[2]
                                new_row['TOTAL'] = total[2]
                                new_row['PAGO'] = pago
                                self.df_general_data = self.df_general_data.append(new_row, ignore_index=True)

                        tbts = pgs.extract_tables()
                        for y in range(len(tbts)):
                            tbt = tbts[y]
                            df_2 = pd.DataFrame(tbt)
                            cantidad_columna = len(df_2.columns)
                            if cantidad_columna == 10:
                                for index, row in enumerate(tbt):
                                    val = tbt[index][0]
                                    if val == None:
                                        break
                                    elif index == 0:
                                        print('')
                                    else:
                                        data.append(row)
                    cabecera_detalle = ['Cod. Principal', 'Cod. Auxiliar', 'Cantidad', 'Descripción', 'Detalle Adicional',
                                        'Precio_Unitario', 'Subsidio', 'Precio_sin_Subsidio', 'Descuento', 'Precio_Total']
                    df = pd.DataFrame(data, columns=cabecera_detalle)
                    df['Cantidad'] = pd.to_numeric(df.Cantidad)
                    df['Subsidio'] = pd.to_numeric(df.Subsidio)
                    df['Precio_sin_Subsidio'] = pd.to_numeric(df.Precio_sin_Subsidio)
                    df['Descuento'] = pd.to_numeric(df.Descuento)
                    df['Precio_Unitario'] = pd.to_numeric(df.Precio_Unitario)
                    df['Precio_Total'] = pd.to_numeric(df.Precio_Total)

                    self.df_general_data['SUBTOTAL_12'] = pd.to_numeric(self.df_general_data.SUBTOTAL_12)
                    self.df_general_data['SUBTOTAL_0'] = pd.to_numeric(self.df_general_data.SUBTOTAL_0)
                    self.df_general_data['TOTAL'] = pd.to_numeric(self.df_general_data.TOTAL)
                    print('Factura: '+ str(file))
                    sheet_name = str(secuencial[0])
                    df.to_excel(self.writer, index=False, sheet_name=sheet_name, startrow=0)
                self.df_general_data.to_excel(self.writer,index=False, sheet_name='Resumen')
        self.writer.save()
        print('********************---Reporte Generado Correctamente----***************************')

def carpeta():
    directorio = filedialog.askdirectory(title='Selecciona la carpeta')
    if directorio != '':
        os.chdir(directorio)
        path = (os.getcwd()).replace('\\', '/')+'/'
        x = extractor(path)
        x.bucle()

ventana = tkinter.Tk()
ventana.geometry("400x250")
label = tkinter.Label(ventana, text="Seleccionar la carpeta donde se encuentran las \n"
                                    "facturas en formato pdf \n\n").pack()
btn_seleccionar = tkinter.Button(ventana, text="Seleccionar", command=carpeta).pack()
contact = tkinter.Label(ventana, text="\n\n\nSoporte \n mcarrion@gbc.com.ec \n 0991557495").pack()
ventana.mainloop()
