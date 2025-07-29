import pyautogui
import pyperclip
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import time

# === CONFIGURAÇÕES ===
caminho_excel = r"C:\Users\ldeliberal\Desktop\NFP_SCRIPT\01DigitacaoExterna.xlsx"
coluna_chaves = 'A'
linha_inicial = 2
x_campo_chave = 788  # Coordenada X do campo "Chave-de-acesso"
y_campo_chave = 506  # Coordenada Y do campo "Chave-de-acesso"
x_botao_salvar = 893  # Coordenada X do botão "Salvar Nota"
y_botao_salvar = 575  # Coordenada Y do botão "Salvar Nota"

# === TIMER INICIAL ===
print("O script começará em 15 segundos. Prepare a tela em modo 'tela cheia'...")
time.sleep(15)

# === Estilo para pintar célula de verde ===
fill_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Verde claro

# === 1. Carregar planilha ===
wb = load_workbook(caminho_excel)
planilha = wb.active

linha = linha_inicial

while True:
    # === 2. Ler a próxima chave ===
    celula = f"{coluna_chaves}{linha}"
    chave = str(planilha[celula].value)

    if chave == "None" or chave.strip() == "":
        print("Fim das chaves ou célula vazia. Encerrando...")
        break

    # === 3. Copiar chave para área de transferência ===
    pyperclip.copy(chave)

    # === 4. Clicar no campo "Chave-de-acesso" para garantir foco ===
    pyautogui.click(x=x_campo_chave, y=y_campo_chave)
    time.sleep(0.5)

    # === 5. Limpar campo anterior ===
    pyautogui.hotkey("ctrl", "a")
    pyautogui.press("delete")
    time.sleep(0.5)

    # === 6. Colar nova chave ===
    pyautogui.hotkey("ctrl", "v")
    print(f"Chave colada: {chave}")

    # === 7. Clicar no botão "Salvar Nota" ===
    time.sleep(1)
    pyautogui.click(x=x_botao_salvar, y=y_botao_salvar)
    print("Botão 'Salvar Nota' clicado.")

    # === 8. Marcar a célula como processada (verde) e salvar ===
    planilha[celula].fill = fill_verde
    wb.save(caminho_excel)

    # === 9. Esperar 7 segundos antes da próxima ===
    time.sleep(5)
    linha += 1
