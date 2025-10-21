"""
informe de pagos por departamento
"""

#!/usr/bin/env python
# coding: utf-8

# In[18]:


import pandas as pd
import os

# Cargar el archivo Excel con los datos
ruta_archivo = r"C:\Users\HP\Desktop\Informe de saldos por departamento.xlsx"  # Cambia por la ruta de tu archivo
df = pd.read_excel(ruta_archivo, sheet_name="INFORME BANCARIO")

# Lista completa de departamentos que debe estar en el resultado
departamentos = ['T1', 'T2', 'T3', 'T4', 'T5', '1A', '1B', '1C', '1D', '1E', '2A', '2B', '2C', '2D', '2E',
                 '3A', '3B', '3C', '3D', '3E', '4A', '4B', '4C', '4D', '4E', '5A', '5B', '5C', '5D', '5E',
                 '6A', '6B', '6C', '6D', '6E', '7A', '7B', '7C', '7D', '7E', '8A', '8B', '8C', '8D', '8E',
                 '9A', '9B', '9C', '9D', '9E']

# Crear un DataFrame con los departamentos
df_departamentos = pd.DataFrame(departamentos, columns=['DEPARTAMENTO/TIENDA'])

# Filtrar los depósitos (considerando que los montos negativos son retiros, así que solo nos interesan los positivos)
df_depositos = df[df['Monto'] > 0]

# Agrupar los montos por departamento y calcular el total por cada uno
depositos_por_departamento = df_depositos.groupby('DEPARTAMENTO/TIENDA')['Monto'].sum().reset_index()

# Hacer un merge con la lista completa de departamentos para asegurar que todos aparezcan, incluso los que no tienen depósitos
resultado_final = pd.merge(df_departamentos, depositos_por_departamento, on='DEPARTAMENTO/TIENDA', how='left')

# Rellenar los NaN (es decir, los departamentos sin depósitos) con 0
resultado_final['Monto'].fillna(0, inplace=True)

# Mostrar el resultado
print(resultado_final)

# Guardar el resultado en un archivo Excel en la misma carpeta que el archivo original
directorio_archivo = os.path.dirname(ruta_archivo)
nombre_salida = os.path.join(directorio_archivo, 'depositos_por_departamento_completo.xlsx')
resultado_final.to_excel(nombre_salida, index=False)

print(f"Archivo guardado en: {nombre_salida}")



# In[ ]:




