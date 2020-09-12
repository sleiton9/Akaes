import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep
from os import listdir


ruta = '.'  
nombre_hoja='data'
archivos = listdir(ruta)
fecha='2020-08-01'
mes='Agosto'
columna=3
fila=63
fila_geo=2
columna_geo=1

for i in archivos:  
    eval1 = i.find('reporte_helios-agosto')
    if(eval1 >= 0 ):
        archivo=i

df = pd.read_excel(archivo, nombre_hoja)

df=df[['FECHA_CREACION','TRAKING_SHOP','VALOR_USD','PRECIO','TIPO','ENVIO','PROVINCIA']]
df= df[(df['FECHA_CREACION']>=fecha)]
df['FECHA_CREACION'] = df['FECHA_CREACION'].map(lambda x: str(x)[0:10])
df['FECHA_CREACION'] = df['FECHA_CREACION'].map(lambda x: str(x)[8:10])
df['TRAKING_SHOP'].replace(np.nan, 'Cancelado', inplace=True)
df['VALOR_USD'].replace(np.nan, 0, inplace=True)

#----------------------CONTEO DE CANCELADOS------------------------------------
df_cancel=df[['FECHA_CREACION','TRAKING_SHOP']]
df_cancel=df_cancel[(df_cancel['TRAKING_SHOP']=='Cancelado')]
df_cancel= df_cancel.groupby(['FECHA_CREACION'], as_index=False).count()
#df_cancel=df_cancel['TRAKING_SHOP'].count()
#-------------------------------------------------------------------------------


#----------------------CONTEO DE TIPO------------------------------------
df_tipo=df[['FECHA_CREACION','TIPO','TRAKING_SHOP']]
df_tipo=df_tipo[(df_tipo['TRAKING_SHOP']!='Cancelado')]
df_tipo= df_tipo.groupby(['FECHA_CREACION','TIPO'], as_index=False).count()
df_tipo=df_tipo.pivot(index='FECHA_CREACION',columns='TIPO') #easy to read, rows=index columns=columns lol
df_tipo[('TRAKING_SHOP', 'manual')].replace(np.nan, 0, inplace=True)
df_tipo[('TRAKING_SHOP', 'MELI')].replace(np.nan, 0, inplace=True)
#-------------------------------------------------------------------------------

#----------------------CONTEO DE ENVIO------------------------------------
df_envio=df[['FECHA_CREACION','ENVIO','TRAKING_SHOP']]
df_envio=df_envio[(df_envio['TRAKING_SHOP']!='Cancelado')]
df_envio= df_envio.groupby(['FECHA_CREACION','ENVIO'], as_index=False).count()
df_envio=df_envio.pivot(index='FECHA_CREACION',columns='ENVIO') #easy to read, rows=index columns=columns lol
df_envio[('TRAKING_SHOP', 'custom')].replace(np.nan, 0, inplace=True)
df_envio[('TRAKING_SHOP', 'me2')].replace(np.nan, 0, inplace=True)
#-------------------------------------------------------------------------------




#----------------------SUMA VALOR_USD y PRECIO(APROBADOS)---------------------------------
df_aprov=df[['FECHA_CREACION','VALOR_USD','PRECIO','TRAKING_SHOP']]
df_aprov=df_aprov[(df_aprov['TRAKING_SHOP']!='Cancelado')]
df_gp_val= df_aprov.groupby(['FECHA_CREACION'], as_index=False).sum()
#-----------------------------------------------------------------------------------------


#--------------------------CONTEO APROBADOS--------------------------------------------------
df_gp_count=df_aprov[['FECHA_CREACION','TRAKING_SHOP']]
df_gp_count= df_gp_count.groupby(['FECHA_CREACION'], as_index=False).count()
#---------------------------------------------------------------------------------------------


#----------------------SUMA VALOR_USD y PRECIO(APROBADOS)---------------------------------
df_geo=df[['PROVINCIA','TRAKING_SHOP']]
df_geo['PROVINCIA'].replace('BOGOTA', 'Bogotá', inplace=True)
df_geo['PROVINCIA'].replace('bogota', 'Bogotá', inplace=True)
df_geo['PROVINCIA'].replace('Bogota', 'Bogotá', inplace=True)
df_geo['PROVINCIA'].replace('Bogotá ', 'Bogotá', inplace=True)
df_geo['PROVINCIA'].replace('Bogotá D.C.', 'Bogotá', inplace=True)
df_geo['PROVINCIA'].replace('MEDELLIN', 'Medellin', inplace=True)
df_geo['PROVINCIA'].replace('medellin', 'Medellin', inplace=True)
df_geo['PROVINCIA'].replace('Medellin ', 'Medellin', inplace=True)
df_geo['PROVINCIA'].replace('NARIÑO', 'Nariño', inplace=True)
df_geo['PROVINCIA'].replace('TOLIMA', 'Tolima', inplace=True)
df_geo['PROVINCIA'].replace('CALDAS', 'Caldas', inplace=True)
df_geo['PROVINCIA'].replace('caldas', 'Caldas', inplace=True)
df_geo['PROVINCIA'].replace('CARTAGENA', 'Cartagena', inplace=True)
df_geo['PROVINCIA'].replace('Bolivar', 'Bolívar', inplace=True)

df_geo=df_geo[(df_geo['TRAKING_SHOP']!='Cancelado')]
df_gp_geo= df_geo.groupby(['PROVINCIA'], as_index=False).count()
#-----------------------------------------------------------------------------------------


#--------------------------------------FECHAS-------------------------------------------
dates = []

df_date= df[(df['FECHA_CREACION']>=fecha)]
df_date = df_date.groupby(['FECHA_CREACION'], as_index=False).sum()
for i in range(0, df_date.shape[0]):
    dates.append(int(df_date.values[i][0]))
#---------------------------------------------------------------------------------------

#-----------------------CONEXION A GOOGLE SHEETS-----------------------------------------------
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]  # integracion google sheets

creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scopes=scopes)  # integracion google sheets
googleClient = gspread.authorize(creds)  # integracion google sheets

worksheet = googleClient.open('Reporte mensual')  # Nombre worksheet

#------------------------------------------------------------------------------------------------



def data_Update(df_gp_count,df_gp_val,df_cancel):
    sheet = worksheet.worksheet('akaesColombia')
    for i in range(0, df_gp_count.shape[0]):
        sheet.update_cell(fila+int(df_gp_count.values[i][0])-1, columna, (df_gp_count.values[i][1]))
        sleep(1.2)

    for i in range(0, df_gp_val.shape[0]):
        sheet.update_cell(fila+int(df_gp_val.values[i][0])-1, columna+1, (df_gp_val.values[i][1]))
        sleep(1.2)

    for i in range(0, df_gp_val.shape[0]):
        sheet.update_cell(fila+int(df_gp_val.values[i][0])-1, columna+2, (df_gp_val.values[i][2]))  
        sleep(1.2)

    for i in range(0, df_cancel.shape[0]):
        sheet.update_cell(fila+int(df_cancel.values[i][0])-1, columna+3, (df_cancel.values[i][1]))  
        sleep(1.2)

    for i in range(0, df_envio.shape[0]):
        sheet.update_cell(fila+int(df_envio.index[i])-1, columna+6, (df_envio.values[i][0]))  
        sleep(1.2)

    for i in range(0, df_envio.shape[0]):
        sheet.update_cell(fila+int(df_envio.index[i])-1, columna+7, (df_envio.values[i][1]))  
        sleep(1.2)

    for i in range(0, df_tipo.shape[0]):
        sheet.update_cell(fila+int(df_tipo.index[i])-1, columna+8, (df_tipo.values[i][0]))  
        sleep(1.2)

    for i in range(0, df_tipo.shape[0]):
        sheet.update_cell(fila+int(df_tipo.index[i])-1, columna+9, (df_tipo.values[i][1]))  
        sleep(1.2)


def geodata_Update(df_gp_geo):
    sheet = worksheet.worksheet('Colombia geo')
    for i in range(0, df_gp_geo.shape[0]):
        sheet.update_cell(fila_geo+i, columna_geo, mes)
        sleep(1.2)
    
    for i in range(0, df_gp_geo.shape[0]):
        sheet.update_cell(fila_geo+i, columna_geo+1, df_gp_geo.values[i][0])
        sleep(1.2)

    for i in range(0, df_gp_geo.shape[0]):
        sheet.update_cell(fila_geo+i, columna_geo+2, df_gp_geo.values[i][1])
        sleep(1.2)



data_Update(df_gp_count,df_gp_val,df_cancel)
#geodata_Update(df_gp_geo)

