import os
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from wordcloud import WordCloud
from sklearn.preprocessing import LabelEncoder
import numpy as np

def create_output_folder():
    output_folder = 'output'
    counter = 1
    while True:
        folder_name = f'{output_folder}_{counter}' if counter > 1 else output_folder
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
            return folder_name
        counter += 1

def plot_ring_chart(data, column_name, output_folder):
    plt.figure(figsize=(12, 8))

    value_counts = data[column_name].value_counts()
    wedges, texts, autotexts = plt.pie(value_counts, startangle=-40, wedgeprops=dict(width=0.4),
                                       autopct=lambda p: '{:.1f}%'.format(p) if p > 0 else '')

    for i, autotext in enumerate(autotexts):
        angle = wedges[i].theta2 - wedges[i].theta1
        theta = (wedges[i].theta1 + wedges[i].theta2) / 2
        r = 1.5 * 0.5 * wedges[i].width  # Aumentar la separación del centro (ajustar el valor de 1.5)
        x = r * np.cos(np.radians(theta))
        y = r * np.sin(np.radians(theta))

        autotext.set_position((x, y))
        autotext.set_color(wedges[i].get_facecolor())
        autotext.set_size(10)  # Reducir el tamaño del número del porcentaje

    plt.legend(wedges, value_counts.index, title=column_name, loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
    plt.title(f'{column_name}', loc='center')  # Ajustar el título

    save_path = os.path.join(output_folder, f'ring_chart_{column_name}.png')
    plt.savefig(save_path, bbox_inches='tight')
    plt.close()

def generate_wordcloud(data, column_name, output_folder):
    plt.figure(figsize=(14, 8))
    plt.subplot(111)

    # Filtrar palabras de 4 o más letras
    text = ' '.join(word for word in data[column_name].astype(str).tolist() if len(word) >= 4)
    
    wordcloud = WordCloud(width=800, height=400, background_color='white', colormap='Blues').generate(text)
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')
    plt.title(f'{column_name}', loc='center')  # Ajustar el título

    save_path = os.path.join(output_folder, f'wordcloud_{column_name}.png')
    plt.savefig(save_path, bbox_inches='tight')
    plt.close()

def correlation_analysis(data, output_folder):
    label_encoder = LabelEncoder()
    for col in data.columns:
        if data[col].dtype == 'object':
            data[col] = label_encoder.fit_transform(data[col].astype(str))

    correlation_matrix = data.corr()

    plt.figure(figsize=(16, 12))
    plt.subplot(211)
    chart = sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', fmt='.2f', linewidths=.5, cbar_kws={"shrink": 0.75})
    chart.set_title('Correlation Matrix - Main')

    save_path = os.path.join(output_folder, 'correlation_matrix_main.png')
    plt.savefig(save_path, bbox_inches='tight')
    plt.close()

    most_correlated = correlation_matrix.unstack().sort_values(ascending=False).drop_duplicates()
    most_correlated = most_correlated[most_correlated != 1.0].head(10)

    plt.figure(figsize=(12, 8))
    plt.subplot(212)
    most_correlated.plot(kind='bar', color='skyblue')
    plt.title('Top 10 Most Correlated Pairs')
    plt.xlabel('Variable Pairs')
    plt.ylabel('Correlation Coefficient')

    save_path = os.path.join(output_folder, 'top_correlations.png')
    plt.savefig(save_path, bbox_inches='tight')
    plt.close()

file_path = input("Por favor, ingresa la ruta del archivo .xlsx: ")

try:
    data = pd.read_excel(file_path)
    output_folder = create_output_folder()
    basic_stats = data.describe(include='all').transpose()
    basic_stats.to_excel(os.path.join(output_folder, 'basic_statistics.xlsx'), float_format='%.2f')

    for column in data.columns:
        if column == 'Comentarios' or column == '¿Qué temas te interesaría aprender sobre fraude?' or column == '¿A qué área perteneces?':
            continue

        if data[column].dtype == 'object':
            plot_ring_chart(data, column, output_folder)
            generate_wordcloud(data, column, output_folder)

    correlation_analysis(data, output_folder)

except FileNotFoundError:
    print("El archivo no pudo ser encontrado. Por favor, verifica la ruta.")
except Exception as e:
    print("Ocurrió un error:", e)
