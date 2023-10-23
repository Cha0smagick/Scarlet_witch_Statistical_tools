# Importar las bibliotecas necesarias
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Función para calcular y guardar estadísticas básicas
def calcular_estadisticas_basicas(df, output_file):
    resumen_estadisticas = df.describe()
    resumen_estadisticas.to_excel(output_file)

# Función para generar y guardar mapas de calor de correlación
def generar_mapas_calor_correlacion(df, output_folder):
    if not os.path.exists(output_folder):  # Corrección: quitar espacio extra
        os.makedirs(output_folder)

    # Verificar y convertir columnas no numéricas
    for columna in df.columns:
        if not pd.api.types.is_numeric_dtype(df[columna]):
            try:
                df[columna] = pd.to_numeric(df[columna], errors='coerce')
            except ValueError:
                pass

    # Gráfico de correlación
    corr = df.corr()

    # Ajustar el tamaño del lienzo
    plt.figure(figsize=(20, 16), dpi=300)  # Ajusta el tamaño del lienzo y alarga el eje y

    # Mapa de calor de correlación general
    plt.figure(figsize=(20, 16), dpi=300)  # Ajusta el tamaño del lienzo y alarga el eje y
    sns.heatmap(corr, annot=True, cmap="coolwarm", fmt=".2f", linewidths=0.5, annot_kws={"size": 6, "ha": 'center', "va": 'center'})  # Ajusta el tamaño de los números y la alineación
    plt.title('Mapa de calor de correlación (General)')
    plt.tight_layout()
    plt.savefig(f'{output_folder}/mapa_calor_correlacion_general.png', dpi=300)
    plt.close()

    # Mapa de calor de correlación de las más significativas
    plt.figure(figsize=(20, 16), dpi=300)  # Ajusta el tamaño del lienzo y alarga el eje y
    sns.heatmap(corr[abs(corr) > 0.6], annot=True, cmap="coolwarm", fmt=".2f", linewidths=0.5, annot_kws={"size": 6, "ha": 'center', "va": 'center'})  # Ajusta el tamaño de los números y la alineación
    plt.title('Mapa de calor de correlación (Significativas)')
    plt.tight_layout()
    plt.savefig(f'{output_folder}/mapa_calor_correlacion_significativas.png', dpi=300)
    plt.close()

    # Mapa de calor de correlación de las más altas
    plt.figure(figsize=(20, 16), dpi=300)  # Ajusta el tamaño del lienzo y alarga el eje y
    sns.heatmap(corr[corr > 0.8], annot=True, cmap="coolwarm", fmt=".2f", linewidths=0.5, annot_kws={"size": 6, "ha": 'center', "va": 'center'})  # Ajusta el tamaño de los números y la alineación
    plt.title('Mapa de calor de correlación (Altas)')
    plt.tight_layout()
    plt.savefig(f'{output_folder}/mapa_calor_correlacion_altas.png', dpi=300)
    plt.close()

    # Mapa de calor de correlación de las más bajas
    plt.figure(figsize=(20, 16), dpi=300)  # Ajusta el tamaño del lienzo y alarga el eje y
    sns.heatmap(corr[corr < -0.6], annot=True, cmap="coolwarm", fmt=".2f", linewidths=0.5, annot_kws={"size": 6, "ha": 'center', "va": 'center'})  # Ajusta el tamaño de los números y la alineación
    plt.title('Mapa de calor de correlación (Bajas)')
    plt.tight_layout()
    plt.savefig(f'{output_folder}/mapa_calor_correlacion_bajas.png', dpi=300)
    plt.close()

# Función principal
def main():
    input_file = input("Por favor, ingrese la ruta del archivo XLSX a analizar: ")
    output_folder = 'output'
    output_stats_file = 'resumen_estadisticas.xlsx'

    # Cargar el archivo xlsx
    try:
        df = pd.read_excel(input_file, header=0)  # Usar la primera fila como encabezado
    except Exception as e:
        print(f"Error al cargar el archivo: {str(e)}")
        return

    # Calcular y guardar estadísticas básicas
    calcular_estadisticas_basicas(df, output_stats_file)

    # Generar y guardar mapas de calor de correlación
    generar_mapas_calor_correlacion(df, output_folder)

if __name__ == "__main__":
    main()
