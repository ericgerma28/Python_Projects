import pyautogui
import time

# Espera 5 segundos antes de realizar o primeiro clique
time.sleep(5)

# Repetir o processo de clique x vezes
for i in range(232):
    print(f"Iteração {i + 1}:")

    # Primeiro clique: na posição específica (2629, 189)
    x, y = 2629, 189
    print(f"Primeiro clique na posição: ({x}, {y})")  # Imprime a posição do primeiro clique (opcional)
    pyautogui.click(x, y)  # Realiza o clique na posição especificada

    # Espera mais 5 segundos antes de realizar o segundo clique
    time.sleep(5)

    # Segundo clique: na posição específica (4740, 189)
    x2, y2 = 4740, 189
    print(f"Segundo clique na posição: ({x2}, {y2})")  # Imprime a posição do segundo clique (opcional)
    pyautogui.click(x2, y2)  # Realiza o clique na posição especificada

    # Espera 5 segundos antes de iniciar a próxima iteração
    time.sleep(5)

print("Processo concluído.")
