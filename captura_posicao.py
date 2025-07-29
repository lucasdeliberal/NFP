import pyautogui
import time

print("Posicione o mouse sobre o botão 'Salvar Nota'.")
print("O programa irá mostrar a posição atual do mouse a cada 2 segundos.")
print("Digite 'sair' e pressione Enter para fechar o programa.\n")

while True:
    x, y = pyautogui.position()
    print(f"Posição atual do mouse: x={x}, y={y}")
    
    # Pausa para não poluir muito o terminal
    time.sleep(2)
    
    # Verifica se o usuário quer sair
    # Como o script fica em loop, usamos input com timeout não é trivial no terminal padrão,
    # então vamos pedir para o usuário pressionar ENTER para continuar ou digitar 'sair' para encerrar.
    print("Pressione ENTER para atualizar a posição, ou digite 'sair' para encerrar.")
    comando = input().strip().lower()
    if comando == 'sair':
        print("Encerrando o programa...")
        break
