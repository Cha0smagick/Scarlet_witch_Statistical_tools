import pandas as pd
import pingouin as pg

# Solicitar la ruta del archivo Excel
ruta_archivo = input("Por favor, ingrese la ruta del archivo Excel: ")

try:
    # Cargar el archivo Excel en un DataFrame
    df = pd.read_excel(ruta_archivo, header=1)  # Cambia header a 1 si la primera fila contiene nombres de columnas

    # Convertir las columnas a números (manejar valores no numéricos o faltantes)
    df = df.apply(pd.to_numeric, errors='coerce')

    # Calcular el coeficiente alfa de Cronbach
    alpha_cronbach = pg.cronbach_alpha(df)

    print("Coeficiente Alfa de Cronbach:", alpha_cronbach)

except Exception as e:
    print("Error al cargar o procesar el archivo Excel:", e)
