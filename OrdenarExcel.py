import pandas as pd

# Leer el archivo omitiendo las primeras 11 filas
df = pd.read_excel("D:\Jescalante\Descargas\Reporte de Asistencias (6).xls",skiprows=11)

# Definimos con la variable nombre_columnas el nombre de las cabeceras de la tabla
nombre_columnas = ['Columna1','IdMarca','Fecha','Columna4','Columna5','Columna6','Columna7','Columna8','Columna9','Columna10','Columna11','Marca1','Columna13','Columna14','Marca2','Columna16','Marca3','Marca4','Columna19','Marca5','Marca6','Id_Dni','Marca7','Columna24','Marca8','Columna26','IdHorario','Columna28']

df.columns = nombre_columnas

# Definir las columnas a eliminar por nombre
columnas_a_eliminar = ['Columna1', 'Columna4', 'Columna5', 'Columna6', 'Columna7', 'Columna8', 'Columna9', 'Columna10', 
                       'Columna11', 'Columna13', 'Columna14', 'Columna16', 'Columna19', 'Columna24', 'Columna26', 
                       'Columna28']

# Eliminar las columnas, drop(): Es el método utilizado para eliminar las columnas. La opción axis=1 indica que estás eliminando columnas (en lugar de filas).
df = df.drop(columnas_a_eliminar, axis=1)

# Definir el nuevo orden de las columnas
nuevo_orden_columnas = ['IdMarca', 'Marca1', 'Marca2', 'Marca3', 'Marca4', 'Marca5', 'Marca6', 'Marca7', 'Marca8', 'Id_Dni', 'IdHorario', 'Fecha']

# Reorganizar las columnas en el DataFrame
df = df[nuevo_orden_columnas]

# Eliminar filas que estén completamente vacías
df = df.dropna(how='all')

# Eliminar el texto "DNI:" y quedarte solo con el número
df['Id_Dni'] = df['Id_Dni'].str.replace('DNI:', '')

# Rellenar las filas vacías de la columna 'Id_DNI' con el último valor no nulo
df['Id_Dni'] = df['Id_Dni'].fillna(method='ffill')

# Identificar las filas que solo tienen el DNI y eliminar esas filas
df = df.dropna(how='all', subset=['Marca1', 'Marca2', 'Marca3', 'Marca4', 'Marca5', 'Marca6', 'Marca7', 'Marca8', 'IdMarca', 'IdHorario', 'Fecha'])

# Verificar que las columnas fueron eliminadas
print(df.head())

# Guardar el DataFrame como un archivo CSV, index=False: Esto asegura que no se guarde el índice (números de fila) en el archivo CSV.
df.to_csv('D:\PYTHON_1\Resultados_excel\prueba_resultado9.csv', index=False)

# Mensaje de confirmación
print("Archivo CSV guardado correctamente.")