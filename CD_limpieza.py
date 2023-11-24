import pandas as pd

#Cargamos el archivo excel en un dataframe
archivo = 'CD_limpieza.xlsx'
df = pd.read_excel(archivo, sheet_name='Hoja1')

#Reemplazamos las comas por n's en las columnas especificadas
df = df.rename( columns={'TIPO_DA„O': 'TIPO_DAÑO', 'DETALLE_DA„O': 'DETALLE_DAÑO',
            'MONTO_TOTAL_DE_ATENCIîN_PRE': 'MONTO_TOTAL_DE_ATENCION_PRE',
            'MONTO_TOTAL_DE_ATENCIîN_CIEN': 'MONTO_TOTAL_DE_ATENCION_CIEN'}) 

#Eliminamos duplicados de acuerdo a la columna CCT
df = df.drop_duplicates(subset=["CCT"], keep='first')
#print(df)

#Realizamos el metodo HOT DECK (K-NN) en las celdas vacias de la columna matricula
for i in range(2, len(df['MATRICULA'])):
    if pd.isnull(df.iloc[i, df.columns.get_loc('MATRICULA')]):
        df.iloc[i, df.columns.get_loc('MATRICULA')] = (df.iloc[i-1, df.columns.get_loc('MATRICULA')] + df.iloc[i-2, df.columns.get_loc('MATRICULA')]) / 2

# Reemplazamos caracteres especiales en la columna 'DETALLE_DAÑO'
df['DETALLE_DAÑO'] = df['DETALLE_DAÑO'].str.replace('„', 'Ñ').str.replace('î', 'O').str.replace('—n', 'on').str.replace('a–','añ').str.replace('ƒ','E').str.replace('ç', 'A')

#Llenar celdas vacías por el tipo de daño verificando la celda de detalle daño 
#funcion determinar_tipo_dano
def determinar_tipo_dano(row):
    if pd.isnull(row['TIPO_DAÑO']):
        if pd.notnull(row['DETALLE_DAÑO']):
            if '1.-REPARACION DE MUROS; 2.-DAÑOS MENORES DE INFRASTRUCTURA;' in row['DETALLE_DAÑO'] or '    Daño Menor' in row['DETALLE_DAÑO']:
                return '3 Menor'
    return row['TIPO_DAÑO']

# Aplica la función a las filas del DataFrame para llenar las celdas vacías en 'TIPO_DAÑO'
df['TIPO_DAÑO'] = df.apply(determinar_tipo_dano, axis=1)

#Calculamos los valores de las columnas COSTO_ORIGINAL_PRE, MONTO_TOTAL_DE_ATENCION_PRE, MONTO_EJERCIDO_PRE

# Llena las celdas vacías de COSTO_ORIGINAL_PRE con los valores de MONTO_TOTAL_DE_ATENCION_PRE O MONTO_EJERCIDO_PRE
df['COSTO_ORIGINAL_PRE'].fillna(df['MONTO_TOTAL_DE_ATENCION_PRE'], inplace=True)
df['COSTO_ORIGINAL_PRE'].fillna(df['MONTO_EJERCIDO_PRE'], inplace=True)

# Llena las celdas vacías de MONTO_TOTAL_DE_ATENCION_PRE con los valores de COSTO_ORIGINAL_PRE O MONTO_EJERCIDO_PRE
df['MONTO_TOTAL_DE_ATENCION_PRE'].fillna(df['COSTO_ORIGINAL_PRE'], inplace=True)
df['MONTO_TOTAL_DE_ATENCION_PRE'].fillna(df['MONTO_EJERCIDO_PRE'], inplace=True)

# Llena las celdas vacías de MONTO_EJERCIDO_PRE con los valores de COSTO_ORIGINAL_PRE O MONTO_TOTAL_DE_ATENCION_PRE
df['MONTO_EJERCIDO_PRE'].fillna(df['COSTO_ORIGINAL_PRE'], inplace=True)
df['MONTO_EJERCIDO_PRE'].fillna(df['MONTO_TOTAL_DE_ATENCION_PRE'], inplace=True)

#Guardar el DataFrame actualizado en un nuevo archivo excel
archivo_nuevo = 'CD_limpieza_nuevo.xlsx'
df.to_excel(archivo_nuevo, index=False)

