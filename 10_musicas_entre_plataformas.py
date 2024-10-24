import pandas as pd
import matplotlib.pyplot as plt

# Carregar o arquivo CSV com os dados das músicas
df = pd.read_csv('Spotify Most Streamed Songs.csv')

# Converter a coluna 'streams' para numérica, removendo possíveis caracteres não numéricos
df['streams'] = pd.to_numeric(df['streams'], errors='coerce')

# Selecionar as 10 músicas mais ouvidas
top_10_songs = df.nlargest(10, 'streams')

# Exibir os dados das 10 músicas mais ouvidas
print("\nTop 10 Músicas Mais Ouvidas:")
print(top_10_songs[['track_name', 'artist(s)_name', 'streams', 
                     'in_spotify_playlists', 'in_apple_playlists', 
                     'in_deezer_playlists']])

# Plotar a comparação de streams por plataforma
plt.figure(figsize=(14, 7))
bar_width = 0.2
index = range(len(top_10_songs))

# Criar barras para cada plataforma
plt.bar(index, top_10_songs['in_spotify_playlists'], bar_width, label='Spotify', color='b')
plt.bar([i + bar_width for i in index], top_10_songs['in_apple_playlists'], bar_width, label='Apple Music', color='g')
plt.bar([i + bar_width * 2 for i in index], top_10_songs['in_deezer_playlists'], bar_width, label='Deezer', color='r')

# Configurações do gráfico
plt.xlabel('Músicas')
plt.ylabel('Número de Playlists')
plt.title('Comparação das 10 Músicas Mais Ouvidas entre Plataformas')
plt.xticks([i + bar_width for i in index], top_10_songs['track_name'], rotation=45)
plt.legend()
plt.tight_layout()
plt.grid(axis='y')
plt.show()
