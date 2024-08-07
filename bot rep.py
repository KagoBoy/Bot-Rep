from pyautogui import *
import pyautogui
from time import sleep
import smtplib, ssl
from openpyxl import Workbook, load_workbook


def click(x, y):
    pyautogui.moveTo(x, y)
    pyautogui.click()

def check_screen(image):
    try:
        button_pos = pyautogui.locateOnScreen(image, confidence=0.7)
        click(button_pos.left, button_pos.top)
        return True
    except pyautogui.ImageNotFoundException:
        return False
        pass
           
    
        
def main():
	i = 0
	planilha = load_workbook("expediente.xlsx")
	aba_ativa = planilha.active
	vetor = []
    
	for celula in aba_ativa["A"]: #Iterando a coluna A da tabela
		vetor.append(celula.value)
    
	while True:
		if(str(vetor[i]) == "stop"):
			print("Todos itens foram reprocessados!")
			break
		sleep(1)
		pyautogui.hotkey('alt', 'u', 'q')
		pyautogui.hotkey('enter')
		pyautogui.typewrite(str(vetor[i]))
		i = i + 1
		print(i, end=" - ")
		sleep(3)
		pyautogui.hotkey('f9')
		pyautogui.hotkey('alt', 's')
		while True:
			if check_screen("ok.png"):
				#arquivo = open("itens reprocessados.txt", "a")
				#arquivo.write("Item: " + vetor[i-1] + " reprocessado! \n")
				#arquivo.close()
				#sleep(0.5)
				print('Item: ' + vetor[i-1] + ' reprocessado!')
				break
			if check_screen("visualizar.png"):
				pyautogui.hotkey("alt", "i")
				pyautogui.hotkey("enter")
				sleep(1.5)
				pyautogui("esc")
				break
			if check_screen("sair.png"):
				break
	
    
                                     
    
main()            
