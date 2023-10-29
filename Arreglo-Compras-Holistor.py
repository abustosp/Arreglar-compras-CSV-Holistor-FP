import pandas as pd
import numpy as np
import os
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showinfo
#!/usr/bin/python3
import tkinter as tk
import tkinter.ttk as ttk

# Leer el txt de 'CUITs a Filtrar.txt' y crear una lista con los CUITs del tipo np.int64
with open('CUITs a Filtrar.txt') as f:
    cuit = f.read().splitlines()
    cuit = [np.int64(i) for i in cuit]

def Abrir_TXT_Filtrar():
    os.startfile("CUITs a Filtrar.txt")

# Definir la función que procesa un archivo
def procesar_archivo(archivo):
    try:
        # Lee el archivo CSV
        df = pd.read_csv(archivo, sep=";", decimal=",", thousands=".", index_col=[0])

        # Realiza las operaciones que necesitas
        #df = pd.read_csv("C:/HolistorBase/comprobantes_202310.csv" , sep=";", decimal="," , thousands=".", index_col=[0])
        df.rename(columns={'Tipo de Comprobante': "Comprobante", 'Punto de Venta': "PV", 'Número de Comprobante': "Numero", 'nueva2':"nueva2", 'nueva':"nueva", 'Tipo Doc. Vendedor': "tipo_doc", 'Nro. Doc. Vendedor': "cuit", 'Denominación Vendedor': "Vendedor", 'Tipo de Cambio': "TC", 'Moneda Original': "Moneda", 'Importe de Per. o Pagos a Cta. de Otros Imp. Nac.': "Perc_Nac", 'Importe No Gravado': "No_Gravado", 'Importe Exento': "Exento", 'Crédito Fiscal Computable': "Credito_Fiscal", 'Importe Total': "Total", 'Importe de Percepciones de Ingresos Brutos': "PIB", 'Importe de Impuestos Municipales': "Muni", 'Importe de Percepciones o Pagos a Cuenta de IVA': "PIVA", 'Importe de Impuestos Internos': "Imp_Int", 'Importe Otros Tributos': "otros", 'Neto Gravado IVA 0%': "Neto_0%", 'Neto Gravado IVA 2,5%': "Neto_2.5%", 'Importe IVA 2,5%': "IVA_2.5%", 'Neto Gravado IVA 5%': "Neto_5%", 'Importe IVA 5%': "IVA_5%", 'Neto Gravado IVA 10,5%': "Neto_10.5%", 'Importe IVA 10,5%': "IVA_10.5%", 'Neto Gravado IVA 21%': "Neto_21%", 'Importe IVA 21%': "IVA_21%", 'Neto Gravado IVA 27%': "Neto_27%", 'Importe IVA 27%': "IVA_27%", 'Total Neto Gravado': 'Total_Neto_Gravado', 'Total IVA': 'Total_IVA'}, inplace=True)
        df.index = pd.to_datetime(df.index, format='%Y-%m-%d').strftime('%d/%m/%Y')
        #df.index = pd.to_datetime(df.index, format='%d/%m/%Y').strftime('%d/%m/%Y')
        #df.index = pd.to_datetime(df.index).strftime('%d/%m/%Y')

        #Borrar fact B (ver porque me borra las A de las misma fecha!)
        Fac_B = [6, 7, 8]
        for i in Fac_B:
            df = df[df["Comprobante"] != i]

        df['cuit'] = df['cuit'].astype(np.int64)    

        # Eliminar los Cuits que aparecen en el txt
        df = df[~df['cuit'].isin(cuit)]

        #crear base para cargar comprobantes por alicuotas
        df.insert(4, "nueva",None), df.insert(5, "nueva2",None)
        col = list(df.columns)
        col[3],col[4],col[5],col[6],col[7],col[8],col[9],col[10],col[11],col[12],col[13],col[14] = col[5],col[4],col[3],col[6],col[7],col[10],col[9],col[14],col[11],col[12], col[13],col[8]
        df = df[col]

        #df.to_excel("C:/HolistorBase/base.xlsx", sheet_name="fact", index_label=None)
        #print(type("Neto_2.5%"))

        #comprobantes al 21%
        df21 = df
        df21 = df21[df21["Comprobante"] != 11]
        df21 = df21[['Comprobante', 'PV', 'Numero', 'nueva2', 'nueva', 'tipo_doc', 'cuit', 'Vendedor', 'TC', 'Moneda', 'Neto_21%', 'No_Gravado', 'Exento', 'IVA_21%', 'Total']]
        df21 = df21[df21['Neto_21%'].notna()]
        #df21.to_excel("C:/HolistorBase/fact_al_21.xlsx", sheet_name="fact", index_label=None)

        #comprobantes al 10.5%
        df1050 = df
        df1050 = df1050[df1050["Comprobante"] != 11]
        df1050 = df1050[['Comprobante', 'PV', 'Numero', 'nueva2', 'nueva', 'tipo_doc', 'cuit', 'Vendedor', 'TC', 'Moneda', 'Neto_10.5%', 'No_Gravado', 'Exento', 'IVA_10.5%', 'Total']]
        df1050 = df1050[df1050['Neto_10.5%'].notna()]
        #df1050.to_excel("C:/HolistorBase/fact_al_105.xlsx", sheet_name="fact", index_label=None)

        #comprobantes al 2.5%
        df250 = df
        df250 = df250[df250["Comprobante"] != 11]
        df250 = df250[['Comprobante', 'PV', 'Numero', 'nueva2', 'nueva', 'tipo_doc', 'cuit', 'Vendedor', 'TC', 'Moneda', 'Neto_2.5%', 'No_Gravado', 'Exento', 'IVA_2.5%', 'Total']]
        df250 = df250[df250['Neto_2.5%'].notna()]
        #df250.to_excel("C:/HolistorBase/fact_al_250.xlsx", sheet_name="fact", index_label=None)

        #comprobantes al 27%
        df27 = df
        df27 = df27[['Comprobante', 'PV', 'Numero', 'nueva2', 'nueva', 'tipo_doc', 'cuit', 'Vendedor', 'TC', 'Moneda', 'Neto_27%', 'No_Gravado', 'Exento', 'IVA_27%', 'Total']]
        df27 = df27[df27['Neto_27%'].notna()]
        #df27.to_excel("C:/HolistorBase/fact_al_27.xlsx", sheet_name="fact", index_label=None)

        #percepción de IB e IVA
        dfib = df
        dfib["CodPIB"] = "PIB"
        dfib = dfib[['Comprobante', 'PV', 'Numero', 'nueva2', 'nueva', 'tipo_doc', 'cuit', 'Vendedor', 'TC', 'Moneda', 'nueva', 'No_Gravado', 'Exento', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'CodPIB', "PIB"]]
        fecha_recep = dfib.index.to_frame(name='fecha')
        dfib.insert(20, 'fecha', fecha_recep)
        dfib.loc[:, 'PIB'] = dfib['PIB'].abs()
        dfib = dfib[dfib['PIB'] >0]
        dfiva = df
        dfiva["CodPIVA"] = "PIVA"
        dfiva = dfiva[['Comprobante', 'PV', 'Numero', 'nueva2', 'nueva', 'tipo_doc', 'cuit', 'Vendedor', 'TC', 'Moneda', 'nueva', 'No_Gravado', 'Exento', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'CodPIVA', "PIVA"]]
        fecha_recep = dfiva.index.to_frame(name='fecha')
        dfiva.insert(20, 'fecha', fecha_recep)
        dfiva.loc[:, 'PIVA'] = dfiva['PIVA'].abs()
        dfiva = dfiva[dfiva['PIVA'] >0]
        dfiva2 = df
        dfiva2["CodPIVA"] = "PIVA"
        dfiva2 = dfiva2[['Comprobante', 'PV', 'Numero', 'nueva2', 'nueva', 'tipo_doc', 'cuit', 'Vendedor', 'TC', 'Moneda', 'nueva', 'No_Gravado', 'Exento', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'nueva', 'CodPIVA', "Perc_Nac"]]
        fecha_recep = dfiva2.index.to_frame(name='fecha')
        dfiva2.insert(20, 'fecha', fecha_recep)
        dfiva2.loc[:, 'Perc_Nac'] = dfiva2['Perc_Nac'].abs()
        dfiva2 = dfiva2[dfiva2['Perc_Nac'] >0]
        dfib.rename(columns={"CodPIB":"Perc", "PIB":"Monto"}, inplace=True)
        dfiva.rename(columns={"CodPIVA":"Perc", "PIVA":"Monto"}, inplace=True)
        dfiva2.rename(columns={"CodPIVA":"Perc", "Perc_Nac":"Monto"}, inplace=True)
        perc = pd.concat([dfib, dfiva,dfiva2])

        #Renonbrar columnas para igualar DF y concatenar
        df21.rename(columns={"Neto_21%":"Neto", "IVA_21%":"Iva"}, inplace=True)
        df1050.rename(columns={"Neto_10.5%":"Neto", "IVA_10.5%":"Iva"}, inplace=True)
        df250.rename(columns={'Neto_2.5%':"Neto", "IVA_2.5%":"Iva"}, inplace=True)
        df27.rename(columns={'Neto_27%':"Neto", "IVA_27%":"Iva"}, inplace=True)

        dffinal = pd.concat([df1050,df21,df250,df27])
        #dffinal['Neto'] = dffinal['Neto'].str.replace('.', ',', regex=False)
        #dffinal['Neto'] = dffinal['Neto'].str.replace(',', '.', regex=False)
        #dffinal['Neto'] = pd.to_numeric(dffinal['Neto'])
        dffinal['Neto'] = dffinal['Neto'].abs()
        dffinal['Iva'] = dffinal['Iva'].abs()
        dffinal['Total'] = dffinal['Total'].abs()
        #print(dffinal.tail(3))


        #Escribir DF en Base de Excel
        #with pd.ExcelWriter('archivo_final.xlsx', engine='xlsxwriter') as writer:
            #dffinal.to_excel(writer, sheet_name='Facturas', index=True)
            #perc.to_excel(writer, sheet_name='Percepciones', index=True)

        # Guarda el DataFrame en un archivo Excel
        archivo_salida = archivo.replace(".csv", "_procesado.xlsx")
        with pd.ExcelWriter(archivo_salida, engine='xlsxwriter') as writer:
            dffinal.to_excel(writer, sheet_name='Facturas', index=True)
            perc.to_excel(writer, sheet_name='Percepciones', index=True)

        print(f"Archivo procesado y guardado como {archivo_salida}")

    except Exception as e:
        print(f"Error al procesar {archivo}: {str(e)}")



def Procesar():
    # Directorio que contiene los archivos CSV
    directorio = askdirectory(title="Selecciona el directorio que contiene los archivos CSV")

    # Obtener la lista de archivos en el directorio
    archivos = [os.path.join(directorio, archivo) for archivo in os.listdir(directorio) if archivo.endswith('.csv')]

    # Iterar sobre la lista de archivos y aplicar la función a cada uno
    for archivo in archivos:
        procesar_archivo(archivo)
    
    showinfo("Proceso finalizado", "Los archivos se han procesado")


class GuiApp:
    def __init__(self, master=None):
        # build ui
        Toplevel_1 = tk.Tk() if master is None else tk.Toplevel(master)
        Toplevel_1.configure(
            background="#ffffff",
            cursor="arrow",
            height=250,
            width=300)
        Toplevel_1.iconbitmap("BIN/ABP-blanco-en-fondo-negro.ico")
        Toplevel_1.overrideredirect("False")
        Toplevel_1.resizable(False, False)
        Toplevel_1.title("Corrección de Compras")
        Label_1 = ttk.Label(Toplevel_1)
        Label_1.configure(
            background="#ffffff",
            cursor="arrow",
            justify="center",
            takefocus=False,
            text='\n\nCorreción masiva de Compras en base el CSV del Portal IVA para la importación en el Holistor\n\n',
            wraplength=300)
        Label_1.pack(expand=True, side="top")
        Label_2 = ttk.Label(Toplevel_1)
        Label_2.configure(
            background="#ffffff",
            justify="center",
            text='Desarrollado por: Federico Perret\nCompilado por: Agustín Bustos Piasentini\nhttps://www.agustin-bustos-piasentini.com.ar/')
        Label_2.pack(expand=True, side="top")
        self.Abir_carpeta_CSV = ttk.Button(Toplevel_1)
        self.Abir_carpeta_CSV.configure(text='Seleccionar carpeta con CSV' , command=Procesar)
        self.Abir_carpeta_CSV.pack(expand=True, pady=4, side="top")
        self.Abrir_TXT_Filtrar = ttk.Button(Toplevel_1)
        self.Abrir_TXT_Filtrar.configure(text='Abrir TXT con CUITs a Filtrar' , command=Abrir_TXT_Filtrar)
        self.Abrir_TXT_Filtrar.pack(expand=True, pady=4, side="top")
        label1 = ttk.Label(Toplevel_1)
        self.img_LogoEstudioPerret = tk.PhotoImage(
            file="BIN/Logo Estudio Perret.png")
        label1.configure(image=self.img_LogoEstudioPerret, text='label1')
        label1.pack(side="top")

        # Main widget
        self.mainwindow = Toplevel_1

    def run(self):
        self.mainwindow.mainloop()


if __name__ == "__main__":
    app = GuiApp()
    app.run()
