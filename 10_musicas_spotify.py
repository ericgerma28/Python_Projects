import pandas as pd
import matplotlib.pyplot as plt

# Carregar o arquivo CSV
df = pd.read_csv('//content//Spotify Most Streamed Songs.csv')

# Converter a coluna 'streams' para numérica, removendo possíveis caracteres não numéricos
df['streams'] = pd.to_numeric(df['streams'], errors='coerce')

# Separar os artistas na coluna 'artist(s)_name' (pois alguns têm colaborações com múltiplos artistas)
df['artists_list'] = df['artist(s)_name'].str.split(', ')

# Explodir a lista de artistas para que cada artista tenha sua própria linha
df_exploded = df.explode('artists_list')

# Agrupar por artista e somar o número de streams
top_10_artists = df_exploded.groupby('artists_list')['streams'].sum().nlargest(10)

# Exibir os 10 artistas mais ouvidos
print("\nTop 10 Artistas Mais Ouvidos:")
print(top_10_artists)

# Plotar os 10 artistas mais ouvidos
plt.figure(figsize=(12, 6))
plt.barh(top_10_artists.index, top_10_artists.values, color='m')
plt.xlabel('Número de Streams')
plt.ylabel('Artista')
plt.title('Top 10 Artistas Mais Ouvidos')
plt.gca().invert_yaxis()  # Inverter o eixo Y para que o artista mais ouvido apareça no topo
plt.grid(True)
plt.show()
