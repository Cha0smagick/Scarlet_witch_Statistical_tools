import pandas as pd
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# Pedir la ruta del archivo de Excel por consola
excel_path = input("Por favor, ingrese la ruta del archivo de Excel: ")

# Leer el archivo de Excel
df = pd.read_excel(excel_path)

# Definir una función para contar palabras en una celda
def contar_palabras(celda):
    palabras = celda.split()
    return Counter(palabras)

# Inicializar un contador global de palabras
conteo_global = Counter()

# Iterar sobre las columnas del DataFrame
for columna in df.columns:
    # Iterar sobre cada celda de la columna
    for celda in df[columna].dropna():
        # Contar palabras en la celda y actualizar el conteo global
        conteo_celda = contar_palabras(str(celda))
        conteo_global.update(conteo_celda)

# Calcular el porcentaje para cada palabra
total_palabras = sum(conteo_global.values())
porcentajes = {palabra: (conteo, conteo / total_palabras * 100) for palabra, conteo in conteo_global.items()}

# Crear un DataFrame con los resultados
resultados_df = pd.DataFrame(list(porcentajes.items()), columns=['Palabra', 'Conteo'])
resultados_df[['Conteo', 'Porcentaje']] = pd.DataFrame(resultados_df['Conteo'].tolist(), index=resultados_df.index)

# Filtrar palabras con 5 o más letras para la nube de palabras
palabras_nube = {palabra: conteo for palabra, conteo in conteo_global.items() if len(palabra) >= 6}

# Crear la nube de palabras
wordcloud = WordCloud(width=800, height=400, background_color='white', colormap='viridis').generate_from_frequencies(palabras_nube)

# Configurar la figura
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud, interpolation='bilinear')
plt.axis('off')

# Guardar la nube de palabras en formato png
wordcloud_path = "nube_palabras.png"
plt.savefig(wordcloud_path, bbox_inches='tight')
plt.close()

# Guardar los resultados en un archivo Excel
resultados_path = "resultados.xlsx"
resultados_df.to_excel(resultados_path, index=False)

print(f"Resultados guardados en {resultados_path}")
print(f"Nube de palabras guardada en {wordcloud_path}")
