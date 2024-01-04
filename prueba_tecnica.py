import pandas as pd
from zipfile import ZipFile
import os

# Ruta del archivo zip descargado
zip_path = 'victims.zip'

# Nombre del archivo Excel de interés (ajusta según sea necesario)
excel_file_name = 'Victims_Age_by_Offense_Category_2022.xlsx'

# Extraer el contenido del archivo zip
with ZipFile(zip_path, 'r') as zip_ref:
    # Verificar si el archivo Excel de interés está presente en el zip
    if excel_file_name not in zip_ref.namelist():
        print(f"El archivo {excel_file_name} no se encuentra en el zip.")
    else:
        # Extraer el archivo Excel
        zip_ref.extract(excel_file_name, path='temp_folder')

        # Leer el archivo Excel con Pandas
        excel_path = os.path.join('temp_folder', excel_file_name)

        # Leer el archivo Excel y obtener los nombres de las columnas
        data = pd.read_excel(excel_path, header=None, skiprows=2)

        # Limpiar los nombres de las columnas
        cleaned_columns = [str(col).replace(' ', '_').replace(',', '') for col in data.iloc[2]]
        data.columns = cleaned_columns

        # Buscar la fila que contiene "Crimes Against Property"
        category_name = "Crimes Against Property"
        category_index = data[data.iloc[:, 0].astype(str).str.contains(category_name)].index

        if not category_index.empty:
            # Excluir el primer valor en la fila que contiene "Crimes Against Property"
            crimes_data = data.iloc[category_index[0]:category_index[0]+1, 1:].apply(lambda x: x.iloc[1:] if x.name == category_index[0] else x, axis=1)

            # Imprimir el conjunto de datos filtrado
            print(f"Conjunto de datos filtrado para '{category_name}':")
            print(crimes_data)

            # Guardar el conjunto de datos filtrado en un archivo CSV sin índice
            csv_file_name = 'Crimes_Against_Property_Data.csv'
            crimes_data.to_csv(csv_file_name, index=False)  # Ajuste: index=False
            print(f"Archivo CSV '{csv_file_name}' generado con éxito.")
        else:
            print(f"No se encontró la categoría '{category_name}'.")
