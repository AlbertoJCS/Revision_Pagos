import time

import numpy as np
import pandas as pa

ruta = r"D:\PAGOS\REVISION DE PAGOS.xlsx"

df = pa.read_excel(ruta,sheet_name="Reporte Pagos Sicsa y Maya", skiprows=4)

#print(df.columns)

# 1- Eliminar columnas innecesarias
df = df.drop(['Unnamed: 0', 'Mes', 'Comentario', 'Unnamed: 12'], axis=1)

print("---Columnas Innecesarias Eliminadas---\n")

# 2- Renombrar el encabezado de las columnas
df.rename(columns={"Nombre": "NOMBRE", "Tarjeta": "TARJETA", "Fecha Real": "FECHA REAL", "Fecha Valor": "FECHA VALOR", "Monto ¢": "MONTO ¢", "Monto $": "MONTO $", "# Data": "OPERACION", "Detalle": "OBSERVACION",  "OCE": "MONEDA"}, inplace=True)

print("---Renombración de columnas---\n")

#print(df.columns)

# 3- Filtrar los registros que coinciden con el criterio "bcr" en la columna "I"
filtered_df = df[df["OBSERVACION"] == "bcr"]

# 4- Obtener los índices de las filas filtradas
indices_to_drop = filtered_df.index

# 5- Eliminar las filas correspondientes utilizando drop()
df = df.drop(indices_to_drop)

print("---Eliminación de Registro del BCR---\n")

# 6- Eliminar el contenido de las columnas de OBSERVACION Y MONEDA
df[["OBSERVACION", "MONEDA"]] = ""

print("---Eliminación de Contenido en la Observación y Moneda---\n")

print(filtered_df.columns)
#time.sleep(30)

# 7- Filtrar los registros de mayor a menor en la columna MONTO ¢ y con los valores ("-") y luego asignar los valores de la columna MONTO $ despues de aplicar el filtro
filtered_df = df[df["MONTO ¢"] == 0]
#filtered_df = filtered_df.sort_values(by="MONTO ¢", ascending=False)
montosDolares = filtered_df["MONTO $"]

df.loc[df["MONTO ¢"] == 0, "MONTO ¢"] = montosDolares
df.loc[df["MONTO ¢"] == 0, "MONEDA"] = "DOLARES"
print("---Mover los Pagos de Dolares---\n")

# 8 - Indicar si es un pago en colones o dolares
df.loc[df["MONEDA"] == "", "MONEDA"] = "COLONES"

print("---Clasificación de Pagos en Dolares Y Colones---\n")

#9 - Eliminar la columna que ya no se va necesitar
df = df.drop("MONTO $", axis=1)

print("---Eliminar Columna de Montos en Dolares---\n")

#10 - Ordenar los registros por el monto de colones de menor a mayor
df.sort_values(by="MONTO ¢", ascending = True)
time.sleep(10)
#11 - Crear el archivo Excel
ruta2 = r"D:\PAGOS"
NombreArchivo = "PAGOS COM"
df.to_excel(ruta2 + "\\" + NombreArchivo + ".xlsx", sheet_name = NombreArchivo)

print("---Excel Creado!---\n")

