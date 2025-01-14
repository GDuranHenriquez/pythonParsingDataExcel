from createWord import create_word_document
import pandas as pd
import os

current_directory = os.path.dirname(__file__)
file_path = os.path.join(current_directory, '..', 'data' , 'controldeinstacionestotal1.xls')

def get_all_name_column(df, name):  
  return df[name].unique().tolist()

if not os.path.exists(file_path):
    print(f"El archivo no existe: {file_path}")
else:
    # Leer el archivo de Excel
    try:
        select_columns = ["nombretitular", "nombreedificacion", "count_fat64", "count_fat32", "count_fat16", "count_fat8", "X", "y", 
                          "tipo_edificacion", "uso_inmueble", "direccion_inm", "parroquia_1", "cod_manz", "cod_parc", "cod_campo_edif", 
                          "cod_estado", "cod_municipio", "cod_parroquia", "cod_ambito", "cod_sector", "sectores"]
        
        df = pd.read_excel(file_path, engine='xlrd', usecols=select_columns)
        print("Archivo leído exitosamente.")
        #print(df.head())
        df = df[df['sectores'].str.lower() == ('Colinas de Bello Monte').lower()]
        unique_colums_edificacion = get_all_name_column(df, 'nombreedificacion')
        
        for name in unique_colums_edificacion:        
            df_filtered  = df[df['nombreedificacion'].str.lower() == name.lower()]
            columns_fat_to_check = ["count_fat64", "count_fat32", "count_fat16", "count_fat8"]
            
            create_doc = False
            if ((df_filtered[columns_fat_to_check] > 0).any().any()):
                create_doc = True
            if(create_doc):
                primera_fila = df_filtered.iloc[0]
                create_word_document(primera_fila, name)
            
        
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
